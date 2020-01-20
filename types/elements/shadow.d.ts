export interface ShadowOptions {
    type?: 'outer' | 'inner' | 'none';
    blur?: number;
    offset?: number;
    angle?: number;
    color?: string;
    opacity?: number | string;
}
export default class ShadowElement {
    type: 'outer' | 'inner' | 'none';
    blur: number;
    offset: number;
    angle: number;
    color: string;
    opacity: number;
    constructor(options: ShadowOptions);
    render(): string;
}
