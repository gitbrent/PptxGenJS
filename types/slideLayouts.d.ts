import Relations from './relations';
import ElementInterface from './elements/element-interface';
import PlaceholderTextElement from './elements/placeholder-text';
import PlaceholderImageElement from './elements/placeholder-image';
declare type Placeholder = PlaceholderImageElement & PlaceholderTextElement;
export declare class Master {
    name: any;
    data: ElementInterface[];
    margin: [number, number, number, number];
    relations: Relations;
    presLayout: any;
    bkgd?: string;
    bkgdImgRid?: number;
    placeholders: Map<string, Placeholder>;
    constructor(title: string, layout: any);
    readonly rels: any[];
    readonly relsChart: any[];
    readonly relsMedia: any[];
    configureBackground(bkg: any): void;
    fromConfig(slideDef: any): void;
    getPlaceholder(placeholderName?: string): PlaceholderTextElement;
}
export default class SlideLayouts {
    private layoutsOrder;
    private layouts;
    private presLayout;
    masterSlide: Master;
    constructor(presLayout: any);
    add(layoutId: any, newLayout: any): void;
    get(layoutId: any): Master;
    provide(layoutId: any): Master;
    new(name: string): Master;
    newFromConfig(name: string, config: any): Master;
    asList(): Master[];
    forEach(arg1: any, arg2?: any): void;
    map(arg1: any, arg2: any): {}[];
    filter(arg1: any, arg2: any): Master[];
}
export {};
