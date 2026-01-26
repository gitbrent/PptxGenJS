/**
 * PptxGenJS: Table Generation
 */
import { PresLayout, SlideLayout, TableCell, TableToSlidesProps, TableRowSlide } from './core-interfaces';
import PptxGenJS from './pptxgen';
/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {TableCell[][]} tableRows - table rows
 * @param {TableToSlidesProps} tableProps - table2slides properties
 * @param {PresLayout} presLayout - presentation layout
 * @param {SlideLayout} masterSlide - master slide
 * @return {TableRowSlide[]} array of table rows
 */
export declare function getSlidesForTableRows(tableRows: TableCell[][], tableProps: TableToSlidesProps, presLayout: PresLayout, masterSlide?: SlideLayout): TableRowSlide[];
/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {PptxGenJS} pptx - pptxgenjs instance
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {ITableToSlidesOpts} options - array of options (e.g.: tabsize)
 * @param {SlideLayout} masterSlide - masterSlide
 */
export declare function genTableToSlides(pptx: PptxGenJS, tabEleId: string, options?: TableToSlidesProps, masterSlide?: SlideLayout): void;
