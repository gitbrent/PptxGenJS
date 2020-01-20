import { TEXT_HALIGN } from '../core-enums';
import Bullet, { BulletOptions } from './bullet';
export interface ParagraphPropertiesOptions {
    rtlMode?: boolean;
    paraSpaceBefore?: number | string;
    paraSpaceAfter?: number | string;
    indentLevel?: number | string;
    bullet?: BulletOptions;
    align?: string;
    lineSpacing?: number;
}
export default class ParagraphProperties {
    bullet: Bullet;
    align: TEXT_HALIGN;
    lineSpacing?: number;
    indentLevel?: number;
    paraSpaceBefore?: number;
    paraSpaceAfter?: number;
    rtlMode?: boolean;
    constructor({ rtlMode, paraSpaceBefore, paraSpaceAfter, indentLevel, bullet, align, lineSpacing }: {
        rtlMode: any;
        paraSpaceBefore: any;
        paraSpaceAfter: any;
        indentLevel: any;
        bullet: any;
        align: any;
        lineSpacing: any;
    });
    render(presLayout: any, tag: string, body?: string): string;
}
