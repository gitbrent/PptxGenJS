import ElementInterface from './element-interface';
import ShadowElement, { ShadowOptions } from './shadow';
import Shape, { ShapeOptions as SO, ShapeConfig } from './shape';
import Position, { PositionOptions } from './position';
import Line from './line';
declare type FullColor = string | {
    type: string;
    color: string;
    alpha?: number;
};
export declare type ShapeOptions = PositionOptions & SO & {
    fill?: FullColor;
    color?: string;
    rectRadius?: number;
    line?: string;
    lineSize?: number;
    lineDash?: string;
    lineHead?: string;
    lineTail?: string;
    shadow?: ShadowOptions;
};
export default class SimpleShapeElement implements ElementInterface {
    shape: Shape;
    fill?: FullColor;
    color?: string;
    position: Position;
    line?: Line;
    shadow?: ShadowElement;
    constructor(shape: ShapeConfig, opts: ShapeOptions);
    render(idx: any, presLayout: any): string;
}
export {};
