import Placeholder, { PlaceholderOptions } from './placeholder';
import { ObjectFitOptions, ColorBlend } from './image';
import Position, { PositionOptions } from './position';
export declare type PlaceholderImageOptions = PositionOptions & PlaceholderOptions & {
    objectFit?: ObjectFitOptions;
    colorBlend?: ColorBlend;
    opacity?: number;
};
export default class PlaceholderImage extends Placeholder {
    position: Position;
    objectFit?: ObjectFitOptions;
    opacity?: number;
    colorBlend?: ColorBlend;
    constructor(options: PlaceholderImageOptions, index: any);
    render(idx: any, presLayout: any): string;
}
