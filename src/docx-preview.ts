import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    useMathMLPolyfill: boolean;
    renderChanges: boolean;
    callback: (e: HTMLElement, text: string) => void
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    useBase64URL: false,
    useMathMLPolyfill: false,
    renderChanges: false,
    callback (e: HTMLElement, text: string) {
        if (text === '<---type:key--->') {
            e.innerText = '狗蛋'
        }
    }
}

export function praseAsync (data: Blob | any, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export function renderAsync (data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(window.document);

    return WordDocument
        .load(data, new DocumentParser(ops), ops)
        .then(doc => {
            load(doc, bodyContainer, styleContainer, ops);
            return doc;
        });
}


export function load (doc: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Options) {
    bodyContainer.innerHTML = ''
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(window.document);
    renderer.render(doc, bodyContainer, styleContainer, ops);
    const articleList = Array.from(bodyContainer.children).filter(el => {
        return el.tagName === 'DIV'
    }) as HTMLElement[]
    function searchTextDom (element: HTMLElement) {
        if (element.children.length) {
            const els: HTMLElement[] = Array.from(element.children) as HTMLElement[]
            for (let el of els) {
                searchTextDom(el)
            }
        } else {
            if (element.innerText.includes('<---') && element.innerText.includes('--->')) {
                ops.callback(element, element.innerText)
                return
            }
        }
    }
    articleList.forEach(el => {
        searchTextDom(el)
    })
}
