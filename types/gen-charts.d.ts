/**
 * PptxGenJS: Chart Generation
 */
import { ISlideRelChart } from './core-interfaces';
import * as JSZip from 'jszip';
/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
 * @param {ISlideRelChart} chartObject - chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
export declare function createExcelWorksheet(chartObject: ISlideRelChart, zip: JSZip): Promise<any>;
/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
export declare function makeXmlCharts(rel: ISlideRelChart): string;
