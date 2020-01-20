import Relations from '../relations';
export interface HyperLinkOptions {
    url?: string;
    slide?: number;
    tooltip?: string;
}
export default class HyperLink {
    url?: string;
    slide?: number;
    tooltip?: string;
    rId?: number;
    constructor({ url, slide, tooltip }: HyperLinkOptions, relations: Relations);
    render(): string;
}
