import ElementInterface from './element-interface';
export default class ChartElement implements ElementInterface {
    chartId: any;
    position: any;
    constructor(type: any, data: any, opts: any, relations: any);
    render(idx: any, presLayout: any): string;
}
