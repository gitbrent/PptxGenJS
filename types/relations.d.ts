export default class Relations {
    rels: any[];
    relsChart: any[];
    relsMedia: any[];
    registerLink(data: any, target: any): number;
    registerImage({ path, data }: {
        path: any;
        data?: string;
    }, extension: any, fromSvgSize?: boolean): number;
    registerChart(globalId: any, options: any, data: any): number;
    registerMedia({ path, type, extn, data, target }: {
        path: any;
        type: any;
        extn: any;
        data?: any;
        target?: any;
    }): number[];
    render(defaultRels: any): string;
}
