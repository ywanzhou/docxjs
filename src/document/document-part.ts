import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { DocumentElement } from "./document";
import xml from './../parser/xml-parser';

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    body: DocumentElement

    parseXml (root: Element) {
        console.log(root)
        this.body = this._documentParser.parseDocumentFile(root);
    }
}
