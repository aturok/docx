// Example on how to use a template document
// Import from 'docx' rather than '../build' if you install from npm
import * as fs from "fs";
import { Document, ImportDotx, Packer, Paragraph } from "../build";

const importDotx = new ImportDotx();
const filePath = "./demo/dotx/template.dotx";

const LoremIpsum = [
    "Adipisicing reprehenderit voluptate non magna incididunt qui eu. Et magna tempor proident magna irure ea. Consequat consectetur esse laborum laboris. Commodo laborum Lorem magna ut aute eiusmod pariatur non. Labore labore dolor mollit non nulla incididunt labore. Excepteur do qui aliqua cupidatat voluptate aute reprehenderit voluptate nulla minim. Cupidatat nostrud officia dolor tempor sunt irure consequat anim mollit mollit eu laboris voluptate amet.",
    "Dolore amet velit deserunt non irure esse ut elit officia consectetur nostrud. Laborum enim ex tempor fugiat fugiat nisi qui. Reprehenderit fugiat excepteur non nulla fugiat sunt nostrud. Commodo ipsum consequat proident veniam sunt nostrud ipsum proident laboris quis reprehenderit ad quis officia. Eu duis culpa ea commodo sunt nulla non magna culpa occaecat ut in elit. Id esse in do in id nostrud.",
    "Excepteur consequat sunt elit sint in adipisicing sint anim consequat veniam ex et aliqua. Aute do eiusmod anim mollit ea consectetur eu Lorem. Occaecat irure minim dolor est minim eu labore sit. Consectetur tempor occaecat deserunt Lorem non tempor. Voluptate sit ipsum officia duis duis aliquip pariatur ipsum eu ipsum occaecat consequat aute sint. Velit dolore enim excepteur sunt duis fugiat commodo amet do minim.",
    "Sunt aliquip ullamco sunt excepteur irure eiusmod aliqua eu consequat. Non duis occaecat aliqua amet reprehenderit veniam aliqua non non. Cillum incididunt reprehenderit mollit anim magna dolor amet nulla aliqua. Adipisicing elit laboris cupidatat veniam excepteur proident laborum sint irure voluptate eu id minim consectetur.",
    "Est veniam occaecat magna ea voluptate et sunt aliqua ad minim. Aliquip magna commodo mollit adipisicing id sint excepteur anim. Est do ut sit ad occaecat consequat et consectetur tempor do non labore ea pariatur. Adipisicing officia elit aliquip dolore fugiat enim. Consectetur non aute culpa sint consectetur sit reprehenderit nulla fugiat dolor in officia. Laboris commodo enim magna consectetur dolore. Culpa officia id ea aliqua.",
    "Laborum irure cupidatat est occaecat eiusmod incididunt labore consequat aliquip veniam sunt duis veniam proident. Tempor non pariatur labore mollit irure. Mollit magna nisi qui exercitation commodo laboris esse occaecat laborum dolore incididunt tempor Lorem aliquip. Ipsum adipisicing sint eu eu magna.",
    "Voluptate consectetur do qui dolore consequat sunt dolor aute aliquip labore officia quis. Nostrud ut nostrud culpa duis velit anim ullamco. Mollit anim sit dolore magna in anim proident qui eu ad velit. Esse pariatur occaecat laborum velit labore elit id qui dolor exercitation incididunt exercitation amet.",
    "Irure magna amet non adipisicing id ea ullamco officia cupidatat sunt aute anim dolore. Exercitation id qui labore in sint esse nulla anim. Consequat labore dolor labore culpa enim ex deserunt nisi reprehenderit eiusmod amet laborum aute. Nulla eu excepteur tempor anim voluptate aliqua culpa quis mollit magna ea reprehenderit in fugiat. Consectetur nulla quis eiusmod voluptate sint culpa sit. Esse ad culpa sint adipisicing aliquip id pariatur velit labore cillum ipsum. Magna duis aute ut veniam esse adipisicing nulla fugiat ea ad cillum tempor.",
    "Dolor pariatur deserunt pariatur proident proident ex magna ut ut ea dolor nulla dolore exercitation. Et non nulla mollit sit cupidatat ut nisi irure pariatur do qui ullamco aute sint. Cillum tempor fugiat ex est qui sint amet duis excepteur velit amet do. Fugiat ut esse Lorem occaecat tempor non dolor deserunt laboris pariatur quis enim ut irure.",
    "Do incididunt elit magna aliquip eiusmod velit aute. Ullamco duis elit pariatur sunt sint commodo laboris. Cillum cillum ut incididunt ullamco pariatur cupidatat. Aliquip voluptate proident laboris exercitation velit fugiat ex quis est. Sit fugiat id est irure cupidatat adipisicing magna veniam irure cillum cillum ad nulla. Et excepteur et consequat excepteur eiusmod nulla mollit eiusmod ex laboris cillum sit cillum nisi.",
    "Adipisicing reprehenderit voluptate non magna incididunt qui eu. Et magna tempor proident magna irure ea. Consequat consectetur esse laborum laboris. Commodo laborum Lorem magna ut aute eiusmod pariatur non. Labore labore dolor mollit non nulla incididunt labore. Excepteur do qui aliqua cupidatat voluptate aute reprehenderit voluptate nulla minim. Cupidatat nostrud officia dolor tempor sunt irure consequat anim mollit mollit eu laboris voluptate amet.",
    "Dolore amet velit deserunt non irure esse ut elit officia consectetur nostrud. Laborum enim ex tempor fugiat fugiat nisi qui. Reprehenderit fugiat excepteur non nulla fugiat sunt nostrud. Commodo ipsum consequat proident veniam sunt nostrud ipsum proident laboris quis reprehenderit ad quis officia. Eu duis culpa ea commodo sunt nulla non magna culpa occaecat ut in elit. Id esse in do in id nostrud.",
    "Excepteur consequat sunt elit sint in adipisicing sint anim consequat veniam ex et aliqua. Aute do eiusmod anim mollit ea consectetur eu Lorem. Occaecat irure minim dolor est minim eu labore sit. Consectetur tempor occaecat deserunt Lorem non tempor. Voluptate sit ipsum officia duis duis aliquip pariatur ipsum eu ipsum occaecat consequat aute sint. Velit dolore enim excepteur sunt duis fugiat commodo amet do minim.",
    "Sunt aliquip ullamco sunt excepteur irure eiusmod aliqua eu consequat. Non duis occaecat aliqua amet reprehenderit veniam aliqua non non. Cillum incididunt reprehenderit mollit anim magna dolor amet nulla aliqua. Adipisicing elit laboris cupidatat veniam excepteur proident laborum sint irure voluptate eu id minim consectetur.",
    "Est veniam occaecat magna ea voluptate et sunt aliqua ad minim. Aliquip magna commodo mollit adipisicing id sint excepteur anim. Est do ut sit ad occaecat consequat et consectetur tempor do non labore ea pariatur. Adipisicing officia elit aliquip dolore fugiat enim. Consectetur non aute culpa sint consectetur sit reprehenderit nulla fugiat dolor in officia. Laboris commodo enim magna consectetur dolore. Culpa officia id ea aliqua.",
    "Laborum irure cupidatat est occaecat eiusmod incididunt labore consequat aliquip veniam sunt duis veniam proident. Tempor non pariatur labore mollit irure. Mollit magna nisi qui exercitation commodo laboris esse occaecat laborum dolore incididunt tempor Lorem aliquip. Ipsum adipisicing sint eu eu magna.",
    "Voluptate consectetur do qui dolore consequat sunt dolor aute aliquip labore officia quis. Nostrud ut nostrud culpa duis velit anim ullamco. Mollit anim sit dolore magna in anim proident qui eu ad velit. Esse pariatur occaecat laborum velit labore elit id qui dolor exercitation incididunt exercitation amet.",
    "Irure magna amet non adipisicing id ea ullamco officia cupidatat sunt aute anim dolore. Exercitation id qui labore in sint esse nulla anim. Consequat labore dolor labore culpa enim ex deserunt nisi reprehenderit eiusmod amet laborum aute. Nulla eu excepteur tempor anim voluptate aliqua culpa quis mollit magna ea reprehenderit in fugiat. Consectetur nulla quis eiusmod voluptate sint culpa sit. Esse ad culpa sint adipisicing aliquip id pariatur velit labore cillum ipsum. Magna duis aute ut veniam esse adipisicing nulla fugiat ea ad cillum tempor.",
    "Dolor pariatur deserunt pariatur proident proident ex magna ut ut ea dolor nulla dolore exercitation. Et non nulla mollit sit cupidatat ut nisi irure pariatur do qui ullamco aute sint. Cillum tempor fugiat ex est qui sint amet duis excepteur velit amet do. Fugiat ut esse Lorem occaecat tempor non dolor deserunt laboris pariatur quis enim ut irure.",
    "Do incididunt elit magna aliquip eiusmod velit aute. Ullamco duis elit pariatur sunt sint commodo laboris. Cillum cillum ut incididunt ullamco pariatur cupidatat. Aliquip voluptate proident laboris exercitation velit fugiat ex quis est. Sit fugiat id est irure cupidatat adipisicing magna veniam irure cillum cillum ad nulla. Et excepteur et consequat excepteur eiusmod nulla mollit eiusmod ex laboris cillum sit cillum nisi.",
];

fs.readFile(filePath, (err, data) => {
    if (err) {
        throw new Error(`Failed to read file ${filePath}.`);
    }

    importDotx.extract(data).then((templateDocument) => {
        const doc = new Document(
            {
                sections: [
                    {
                        properties: {
                            titlePage: templateDocument.titlePageIsDefined,
                        },
                        children: LoremIpsum.map((text) => new Paragraph(text)),
                    },
                ],
            },
            {
                template: templateDocument,
            },
        );

        Packer.toBuffer(doc).then((buffer) => {
            fs.writeFileSync("My Document.docx", buffer);
        });
    });
});
