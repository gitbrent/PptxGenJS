import Relations from '../relations';
import Hyperlink, { HyperLinkOptions } from './hyperlink';
export interface RunPropertiesOptions {
    lang?: string;
    fontFace?: string;
    fontSize?: number;
    charSpacing?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    underline?: boolean;
    subscript?: boolean;
    superscript?: boolean;
    outline?: {
        size?: number;
        color?: string;
    };
    hyperlink?: HyperLinkOptions;
}
export default class RunProperties {
    lang: string;
    altLang: string;
    fontFace?: string;
    fontSize?: number;
    charSpacing?: number;
    color?: string;
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
    subscript: boolean;
    superscript: boolean;
    outline?: {
        size: number;
        color: string;
    };
    hyperlink?: Hyperlink;
    constructor(options: RunPropertiesOptions, relations: Relations);
    render(tag: any): string;
}
