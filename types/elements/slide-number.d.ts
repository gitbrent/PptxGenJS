import ElementInterface from './element-interface';
export default class SlideNumberElement implements ElementInterface {
    position: any;
    runProperties: any;
    fieldId: any;
    constructor({ x, y, w, h, ...runOptions }: {
        [x: string]: any;
        x: any;
        y: any;
        w: any;
        h: any;
    }, relations: any);
    render(idx: any, presLayout: any, placeholder: any): string;
}
