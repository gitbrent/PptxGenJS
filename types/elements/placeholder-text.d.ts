import Placeholder, { PlaceholderOptions } from './placeholder';
import Relations from '../relations';
import { TextOptions } from './text';
export declare type PlaceholderTextOptions = TextOptions & PlaceholderOptions;
export default class PlaceholderText extends Placeholder {
    private textElement;
    constructor(text: string, options: PlaceholderTextOptions, index: number, relations: Relations);
    readonly position: any;
    renderPlaceholderInfo(): string;
    render(idx: any, presLayout: any): any;
}
