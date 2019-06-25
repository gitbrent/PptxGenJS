/**
 * PptxGenJS: XML Generation
 */
import { ISlide, ISlideRelChart, ITextOpts, ILayout, ISlideLayout, ISlideDataObject } from './interfaces';
export declare var gObjPptxGenerators: {
    /**
     * Adds a background image or color to a slide definition.
     * @param {String|Object} bkg color string or an object with image definition
     * @param {ISlide} target slide object that the background is set to
     */
    addBackgroundDefinition: (bkg: string | {
        src?: string;
        path?: string;
        data?: string;
    }, target: ISlide) => void;
    /**
     * Adds a text object to a slide definition.
     * @param {String} text
     * @param {ITextOpts} opt
     * @param {ISlide} target - slide object that the text should be added to
     * @param {Boolean} isPlaceholder
     * @since: 1.0.0
     */
    addTextDefinition: (text: string | object[], opt: ITextOpts, target: ISlide, isPlaceholder: boolean) => {
        type: any;
        text: any;
        options: any;
    };
    /**
     * Adds Notes to a slide.
     * @param {String} `notes`
     * @param {Object} opt (*unused*)
     * @param {ISlide} `target` slide object
     * @since 2.3.0
     */
    addNotesDefinition: (notes: string, opt: object, target: ISlide) => ISlideDataObject;
    /**
     * Adds a placeholder object to a slide definition.
     * @param {String} `text`
     * @param {Object} `opt`
     * @param {ISlide} `target` slide object that the placeholder should be added to
     */
    addPlaceholderDefinition: (text: string, opt: object, target: ISlide) => {
        type: any;
        text: any;
        options: any;
    };
    /**
     * Adds a shape object to a slide definition.
     * @param {gObjPptxShapes} shape shape const object (pptx.shapes)
     * @param {Object} opt
     * @param {Object} target slide object that the shape should be added to
     * @return {Object} shape object
     */
    addShapeDefinition: (shape: any, opt: any, target: any) => {
        type: any;
        text: any;
        options: {};
    };
    /**
     * Adds an image object to a slide definition.
     * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
     * @param {Object} objImage - object containing `path`/`data`, `x`, `y`, etc.
     * @param {Object} target - slide that the image should be added to (if not specified as the 2nd arg)
     * @return {Object} image object
     */
    addImageDefinition: (objImage: any, target: any) => {
        type: any;
        text: any;
        options: any;
        image: any;
        imageRid: any;
        hyperlink: any;
    };
    /**
     * Generate the chart based on input data.
     * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
     *
     * @param {object} type should belong to: 'column', 'pie'
     * @param {object} data a JSON object with follow the following format
     * @param {object} opt
     * @param {object} target slide object that the chart should be added to
     * @return {Object} chart object
     * {
     *   title: 'eSurvey chart',
     *   data: [
     *		{
     *			name: 'Income',
     *			labels: ['2005', '2006', '2007', '2008', '2009'],
     *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
     *		},
     *		{
     *			name: 'Expense',
     *			labels: ['2005', '2006', '2007', '2008', '2009'],
     *			values: [18.1, 22.8, 23.9, 25.1, 25]
     *		}
     *	 ]
     *	}
     */
    addChartDefinition: (type: any, data: any, opt: any, target: any) => {
        type: any;
        text: any;
        options: any;
        chartRid: any;
    };
    /**
     * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
     * The following object is expected as a slide definition:
     * {
     *   bkgd: 'FF00FF',
     *   objects: [{
     *     text: {
     *       text: 'Hello World',
     *       x: 1,
     *       y: 1
     *     }
     *   }]
     * }
     * @param {Object} slideDef slide definition
     * @param {Object} target empty slide object that should be updated by the passed definition
     */
    createSlideObject: (slideDef: any, target: any) => void;
    /**
     * Transforms a slide object to resulting XML string.
     * @param {ISlide} slideObject slide object created within gObjPptxGenerators.createSlideObject
     * @return {string} XML string with <p:cSld> as the root
     */
    slideObjectToXml: (slideObject: ISlide) => string;
    /**
     * Transforms slide relations to XML string.
     * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
     * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
     *
     * @param {ISlide} slideObject slide object whose relations are being transformed
     * @param {Object[]} defaultRels array of default relations (such objects expected: { target: <filepath>, type: <schemepath> })
     * @return {string} complete XML string ready to be saved as a file
     */
    slideObjectRelationsToXml: (slideObject: ISlide, defaultRels: any) => string;
    imageSizingXml: {
        cover: (imgSize: any, boxDim: any) => string;
        contain: (imgSize: any, boxDim: any) => string;
        crop: (imageSize: any, boxDim: any) => string;
    };
    /**
     * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
     * @param {ISlideRelChart} chartObject chart object
     * @param {jszip} zip file that the resulting XLSX should be added to
     * @return {Promise} promise of generating the XLSX file
     */
    createExcelWorksheet: (chartObject: ISlideRelChart, zip: any) => Promise<any>;
};
/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 */
export declare function makeXmlCharts(rel: ISlideRelChart): string;
/**
* DESC: Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
* EX:
    <p:txBody>
        <a:bodyPr wrap="none" lIns="50800" tIns="50800" rIns="50800" bIns="50800" anchor="ctr">
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:pPr marL="228600" indent="-228600"><a:buSzPct val="100000"/><a:buChar char="&#x2022;"/></a:pPr>
          <a:r>
            <a:t>bullet 1 </a:t>
          </a:r>
          <a:r>
            <a:rPr>
              <a:solidFill><a:srgbClr val="7B2CD6"/></a:solidFill>
            </a:rPr>
            <a:t>colored text</a:t>
          </a:r>
        </a:p>
      </p:txBody>
* NOTES:
* - PPT text lines [lines followed by line-breaks] are createing using <p>-aragraph's
* - Bullets are a paragprah-level formatting device
*
* @param slideObj (object) - slideObj -OR- table `cell` object
* @returns XML string containing the param object's text and formatting
*/
export declare function genXmlTextBody(slideObj: any): string;
export declare function makeXmlContTypes(slides: Array<ISlide>, slideLayouts: any, masterSlide?: any): string;
export declare function makeXmlRootRels(): string;
export declare function makeXmlApp(slides: Array<ISlide>, company: string): string;
export declare function makeXmlCore(title: string, subject: string, author: string, revision: string): string;
export declare function makeXmlPresentationRels(slides: Array<ISlide>): string;
/**
 * Generates XML for the slide file
 * @param {Object} objSlide - the slide object to transform into XML
 * @return {string} strXml - slide OOXML
 */
export declare function makeXmlSlide(objSlide: ISlide): string;
export declare function getNotesFromSlide(objSlide: ISlide): string;
export declare function makeXmlNotesSlide(objSlide: ISlide): string;
/**
 * Generates the XML layout resource from a layout object
 *
 * @param {ISlide} objSlideLayout - slide object that represents layout
 * @return {string} strXml - slide OOXML
 */
export declare function makeXmlLayout(objSlideLayout: ISlideLayout): string;
/**
 * Generates XML for the master file
 * @param {ISlide} objSlide - slide object that represents master slide layout
 * @param {ISlideLayout[]} slideLayouts - slide layouts
 * @return {string} strXml - slide OOXML
 */
export declare function makeXmlMaster(objSlide: ISlide, slideLayouts: Array<ISlideLayout>): string;
/**
 * Generate XML for Notes Master
 *
 * @returns {string} XML
 */
export declare function makeXmlNotesMaster(): string;
/**
 * Generates XML string for a slide layout relation file.
 * @param {Number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
export declare function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: Array<ISlideLayout>): string;
/**
 * Generates XML string for a slide relation file.
 * @param {Number} slideNumber 1-indexed number of a layout that relations are generated for
 * @return {string} complete XML string ready to be saved as a file
 */
export declare function makeXmlSlideRel(slides: Array<ISlide>, slideLayouts: Array<ISlideLayout>, slideNumber: number): string;
/**
 * Generates XML string for a slide relation file.
 * @param {Number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
export declare function makeXmlNotesSlideRel(slideNumber: number): string;
/**
 * Generates XML string for the master file.
 * @param {ISlide} `masterSlideObject` - slide object
 * @return {String} complete XML string ready to be saved as a file
 */
export declare function makeXmlMasterRel(masterSlideObject: ISlide, slideLayouts: Array<ISlideLayout>): string;
export declare function makeXmlNotesMasterRel(): string;
export declare function makeXmlTheme(): string;
/**
* Create the `ppt/presentation.xml` file XML
* @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
* @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
* @param `slides` {Array<ISlide>} presentation slides
* @param `pptLayout` {ISlideLayout} presentation layout
*/
export declare function makeXmlPresentation(slides: Array<ISlide>, pptLayout: ILayout): string;
export declare function makeXmlPresProps(): string;
export declare function makeXmlTableStyles(): string;
export declare function makeXmlViewProps(): string;
export declare function createHyperlinkRels(slides: Array<ISlide>, inText: any, slideRels: any): void;
