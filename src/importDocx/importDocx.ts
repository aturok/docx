import * as fastXmlParser from "fast-xml-parser";
import * as JSZip from "jszip";

import { FooterReferenceType } from "file/document/body/section-properties/footer-reference";
import { HeaderReferenceType } from "file/document/body/section-properties/header-reference";
import { FooterWrapper } from "file/footer-wrapper";
import { HeaderWrapper } from "file/header-wrapper";
import { convertToXmlComponent, ImportedXmlComponent, parseOptions } from "file/xml-components";

const schemeToType = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header": "header",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer": "footer",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image": "image",
};

interface IDocumentRefs {
    headers: Array<{ id: number; type: HeaderReferenceType }>;
    footers: Array<{ id: number; type: FooterReferenceType }>;
}

interface IRelationFileInfo {
    id: number;
    targetFile: string;
    type: "header" | "footer" | "image";
}

type DocumentHeaders = Array<{ type: HeaderReferenceType; header: HeaderWrapper }>;
type DocumentFooters = Array<{ type: FooterReferenceType; footer: FooterWrapper }>;

export interface ITemplateDocument {
    currentRelationshipId: number;
    headers: DocumentHeaders;
    footers: DocumentFooters;
}

export class ImportDocx {
    private currentRelationshipId: number;

    constructor() {
        this.currentRelationshipId = 1;
    }

    public async extract(data: Buffer): Promise<ITemplateDocument> {
        const zipContent = await JSZip.loadAsync(data);

        const documentContent = zipContent.files["word/document.xml"];
        const documentRefs: IDocumentRefs = this.extractDocumentRefs(await documentContent.async("text"));

        const relationshipContent = zipContent.files["word/_rels/document.xml.rels"];
        const documentRelations: IRelationFileInfo[] = this.findReferenceFiles(await relationshipContent.async("text"));

        const headers: DocumentHeaders = [];
        for (const headerRef of documentRefs.headers) {
            const headerKey = "w:hdr";
            const relationFileInfo = documentRelations.find((rel) => rel.id === headerRef.id);
            if (relationFileInfo === null || !relationFileInfo) {
                throw new Error(`Can not find target file for id ${headerRef.id}`);
            }

            const xmlData = await zipContent.files[`word/${relationFileInfo.targetFile}`].async("text");
            const xmlObj = fastXmlParser.parse(xmlData, parseOptions);

            const importedComp = convertToXmlComponent(headerKey, xmlObj[headerKey]) as ImportedXmlComponent;

            const header = new HeaderWrapper(this.currentRelationshipId++, importedComp);
            await this.addImagesToWrapper(relationFileInfo, zipContent, header);
            headers.push({ type: headerRef.type, header });
        }

        const footers: DocumentFooters = [];
        for (const footerRef of documentRefs.footers) {
            const footerKey = "w:ftr";
            const relationFileInfo = documentRelations.find((rel) => rel.id === footerRef.id);
            if (relationFileInfo === null || !relationFileInfo) {
                throw new Error(`Can not find target file for id ${footerRef.id}`);
            }
            const xmlData = await zipContent.files[`word/${relationFileInfo.targetFile}`].async("text");
            const xmlObj = fastXmlParser.parse(xmlData, parseOptions);
            const importedComp = convertToXmlComponent(footerKey, xmlObj[footerKey]) as ImportedXmlComponent;

            const footer = new FooterWrapper(this.currentRelationshipId++, importedComp);
            await this.addImagesToWrapper(relationFileInfo, zipContent, footer);
            footers.push({ type: footerRef.type, footer });
        }

        const templateDocument: ITemplateDocument = { headers, footers, currentRelationshipId: this.currentRelationshipId };
        return templateDocument;
    }

    public async addImagesToWrapper(
        relationFile: IRelationFileInfo,
        zipContent: JSZip,
        wrapper: HeaderWrapper | FooterWrapper,
    ): Promise<void> {
        let wrapperImagesReferences: IRelationFileInfo[] = [];
        const refFile = zipContent.files[`word/_rels/${relationFile.targetFile}.rels`];
        if (refFile) {
            const xmlRef = await refFile.async("text");
            wrapperImagesReferences = this.findReferenceFiles(xmlRef).filter((r) => r.type === "image");
        }
        for (const r of wrapperImagesReferences) {
            const buffer = await zipContent.files[`word/${r.targetFile}`].async("nodebuffer");
            wrapper.addImageRelationship(buffer, r.id);
        }
    }

    public findReferenceFiles(xmlData: string): IRelationFileInfo[] {
        const xmlObj = fastXmlParser.parse(xmlData, parseOptions);
        const relationXmlArray = Array.isArray(xmlObj.Relationships.Relationship)
            ? xmlObj.Relationships.Relationship
            : [xmlObj.Relationships.Relationship];
        const relations: IRelationFileInfo[] = relationXmlArray
            .map((item) => {
                return {
                    id: this.parseRefId(item._attr.Id),
                    type: schemeToType[item._attr.Type],
                    targetFile: item._attr.Target,
                };
            })
            .filter((item) => item.type !== null);
        return relations;
    }

    public extractDocumentRefs(xmlData: string): IDocumentRefs {
        const xmlObj = fastXmlParser.parse(xmlData, parseOptions);
        const sectionProp = xmlObj["w:document"]["w:body"]["w:sectPr"];

        const headersXmlArray = Array.isArray(sectionProp["w:headerReference"])
            ? sectionProp["w:headerReference"]
            : [sectionProp["w:headerReference"]];
        const headers = headersXmlArray.map((item) => {
            return {
                type: item._attr["w:type"],
                id: this.parseRefId(item._attr["r:id"]),
            };
        });

        const footersXmlArray = Array.isArray(sectionProp["w:footerReference"])
            ? sectionProp["w:footerReference"]
            : [sectionProp["w:footerReference"]];
        const footers = footersXmlArray.map((item) => {
            return {
                type: item._attr["w:type"],
                id: this.parseRefId(item._attr["r:id"]),
            };
        });

        return { headers, footers };
    }

    public parseRefId(str: string): number {
        const match = /^rId(\d+)$/.exec(str);
        if (match === null) {
            throw new Error("Invalid ref id");
        }
        return parseInt(match[1], 10);
    }
}
