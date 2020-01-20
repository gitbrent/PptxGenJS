import Position from './position';
import { ILayout, ITableCell, ITableOptions, TableRow } from '../core-interfaces';
import Slide from '../slide';
import { Master } from '../slideLayouts';
import { SLIDE_OBJECT_TYPES } from '../core-enums';
import ElementInterface from './element-interface';
/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragprah-level formatting device
 * @param {ITableCell} slideObj - slideObj -OR- table `cell` object
 * @returns XML containing the param object's text and formatting
 */
export declare function genXmlTextBody(slideObj: ITableCell): string;
export default class TableElement implements ElementInterface {
    type: SLIDE_OBJECT_TYPES;
    arrTabRows: any;
    options: any;
    position: Position;
    constructor(target: Slide, tableRows: TableRow[], options: ITableOptions, slideLayout: Master, presLayout: ILayout, addSlide: Function, getSlide: Function);
    render(idx: any, presLayout: any): string;
}
