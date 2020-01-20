import { IImageOpts, IMediaOpts, IShape, IShapeOptions, IText, ITextOpts } from '../core-interfaces';
import Relations from '../relations';
import ElementInterface from './element-interface';
declare type PositionnedElement = ElementInterface & {
    position: {
        xPos: any;
        yPos: any;
    };
};
export default class GroupElement implements ElementInterface {
    data: PositionnedElement[];
    relations: Relations;
    constructor(relations: Relations);
    readonly position: {
        xPos(presLayout: any): number[];
        yPos(presLayout: any): number[];
    };
    addSlideNumber(value: any): GroupElement;
    addChart(type: any, data: any, options: any): GroupElement;
    addImage(options: IImageOpts): GroupElement;
    addMedia(options: IMediaOpts): GroupElement;
    addShape(shape: IShape, options?: IShapeOptions): GroupElement;
    addText(text: string | IText[], options?: ITextOpts): GroupElement;
    newGroup(): GroupElement;
    render(idx: any, presLayout: any): string;
}
export {};
