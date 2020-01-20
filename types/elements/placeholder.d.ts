import ElementInterface from './element-interface';
import Position from './position';
export declare type PlaceholderOptions = {
    name: string;
    type: string;
};
export default class Placeholder implements ElementInterface {
    name: string;
    position: Position;
    placeholderType: any;
    protected placeholderIndex: any;
    constructor(name: any, type: string, index: any);
    renderPlaceholderInfo(): string;
    render(idx: any, presLayout: any): string;
}
