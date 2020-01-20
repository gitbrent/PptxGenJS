/**
 * PptxGenJS: Table Generation
 */
import PptxGenJS from './pptxgen';
import { ILayout, ITableToSlidesCell, ITableToSlidesOpts, TableRowSlide } from './core-interfaces';
import { Master } from './slideLayouts';
/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {[ITableToSlidesCell[]?]} tableRows - HTMLElementID of the table
 * @param {ITableToSlidesOpts} tabOpts - array of options (e.g.: tabsize)
 * @param {ILayout} presLayout - Presentation layout
 * @param {Master} masterSlide - master slide (if any)
 * @return {TableRowSlide[]} array of table rows
 */
export declare function getSlidesForTableRows(tableRows: [ITableToSlidesCell[]?], tabOpts: ITableToSlidesOpts, presLayout: ILayout, masterSlide: Master): TableRowSlide[];
/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
 */
export declare function genTableToSlides(pptx: PptxGenJS, tabEleId: string, options: ITableToSlidesOpts, masterSlide: Master): void;
