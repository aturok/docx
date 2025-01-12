import { XmlComponent } from "@file/xml-components";
import { LatentStyleException } from "./exceptions";

export class LatentStyles extends XmlComponent {
    constructor(latentException?: LatentStyleException) {
        super("w:latentStyles");

        if (latentException) {
            this.root.push(latentException);
        }
    }
}
