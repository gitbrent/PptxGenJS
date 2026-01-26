/**
 * PptxGenJS: Chart Generation
 */
import { CHART_TYPE } from './core-enums';
import { ISlideRelChart } from './core-interfaces';
import JSZip from 'jszip';
/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
 * @param {ISlideRelChart} chartObject - chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
export declare function createExcelWorksheet(chartObject: ISlideRelChart, zip: JSZip): Promise<string>;
/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
export declare function makeXmlCharts(rel: ISlideRelChart): string;
/**
 * Check if a chart type is a ChartEx type
 * @param {CHART_TYPE} chartType - the chart type to check
 * @return {boolean} true if ChartEx type
 */
export declare function isChartExType(chartType: CHART_TYPE | string): boolean;
/**
 * Get the ChartEx layoutId for a chart type
 * @param {CHART_TYPE} chartType - the chart type
 * @return {string} the layoutId
 */
export declare function getChartExLayoutId(chartType: CHART_TYPE | string): string;
/**
 * Generate ChartEx XML (for treemap, sunburst, histogram, pareto, boxWhisker, waterfall, funnel charts)
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
export declare function makeXmlChartEx(rel: ISlideRelChart): string;
