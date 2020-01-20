import Position from './position';
export declare type ShapeConfig = string | {
    displayName: string;
    name: string;
    avLst: {
        [key: string]: number;
    };
};
export declare type ShapeOptions = {
    rectRadius?: number;
};
export default class ShapeElement {
    displayName: string;
    name: string;
    avLst: string;
    rectRadius?: number;
    constructor(input?: ShapeConfig, options?: ShapeOptions);
    private renderRadius;
    render(position: Position, presLayout: any): string;
}
