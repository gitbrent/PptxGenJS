export interface PositionOptions {
    x?: number | string;
    y?: number | string;
    w?: number | string;
    h?: number | string;
    flipH?: boolean;
    flipV?: boolean;
    rotate?: number;
}
export default class Position {
    x?: number | string;
    y?: number | string;
    w?: number | string;
    h?: number | string;
    flipH?: boolean;
    flipV?: boolean;
    rotate?: number;
    constructor({ x, y, w, h, flipH, flipV, rotate }: PositionOptions);
    cx(presLayout: any): number;
    cy(presLayout: any): number;
    xPos(presLayout: any): number[];
    yPos(presLayout: any): number[];
    render(presLayout: any, tag?: string): string;
}
