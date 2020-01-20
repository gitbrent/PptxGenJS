import { TEXT_VALIGN } from '../core-enums';
export declare type VerticalOptions = 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl';
export interface BodyPropertiesOptions {
    autoFit?: boolean;
    vert?: VerticalOptions;
    shrinkText?: boolean;
    inset?: number;
    margin?: number | [number, number, number, number];
    valign?: string;
    wrap?: string;
}
export default class BodyProperties {
    autoFit: boolean;
    shrinkText: boolean;
    anchor: TEXT_VALIGN;
    vert?: VerticalOptions;
    lIns: number;
    rIns: number;
    tIns: number;
    bIns: number;
    wrap?: string;
    constructor(opts: BodyPropertiesOptions, disableDefaults: boolean);
    render(): string;
}
