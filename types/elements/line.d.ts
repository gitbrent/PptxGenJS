export default class LineElement {
    size: any;
    color: any;
    dash: any;
    head: any;
    tail: any;
    constructor({ size, color, dash, head, tail }: {
        size?: number;
        color?: string;
        dash: any;
        head: any;
        tail: any;
    });
    render(): string;
}
