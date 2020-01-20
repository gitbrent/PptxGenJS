import { ILayout } from './core-interfaces';
import { Master } from './slideLayouts';
import Slide from './slide';
/**
 * Generate XML ContentType
 * @param {Slide[]} slides - slides
 * @param {Master[]} slideLayouts - slide layouts
 * @param {Slide} masterSlide - master slide
 * @returns XML
 */
export declare function makeXmlContTypes(slides: Slide[], slideLayouts: Master[], masterSlide?: Slide): string;
/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export declare function makeXmlRootRels(): string;
/**
 * Creates `docProps/app.xml`
 * @param {Slide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export declare function makeXmlApp(slides: Slide[], company: string): string;
/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} company - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
export declare function makeXmlCore(title: string, subject: string, author: string, revision: string): string;
/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {Slide[]} slides - Presenation Slides
 * @returns XML
 */
export declare function makeXmlPresentationRels(slides: Array<Slide>): string;
/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {Slide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export declare function makeXmlSlide(slide: Slide): string;
/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {Slide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export declare function makeXmlNotesSlide(slide: Slide): string;
/**
 * Generates the XML layout resource from a layout object
 * @param {Master} layout - slide layout (master)
 * @return {string} XML
 */
export declare function makeXmlLayout(layout: Master): string;
/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {Slide} slide - slide object that represents master slide layout
 * @param {Master[]} layouts - slide layouts
 * @return {string} XML
 */
export declare function makeXmlMaster(slide: Slide, layouts: Master[]): string;
/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {Master[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export declare function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: Master[]): string;
/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {Slide[]} slides
 * @param {Master[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export declare function makeXmlSlideRel(slides: Slide[], slideLayouts: Master[], slideNumber: number): string;
/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export declare function makeXmlNotesSlideRel(slideNumber: number): string;
/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {Slide} masterSlide - Slide object
 * @param {Master[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export declare function makeXmlMasterRel(masterSlide: Slide, slideLayouts: Master[]): string;
/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export declare function makeXmlNotesMasterRel(): string;
/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {Slide[]} slides - array of slides
 * @param {ILayout} pptLayout - presentation layout
 * @param {boolean} rtlMode - RTL mode
 * @return {string} XML
 */
export declare function makeXmlPresentation(slides: Slide[], pptLayout: ILayout, rtlMode: boolean): string;
/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export declare function makeXmlPresProps(): string;
/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export declare function makeXmlTableStyles(): string;
/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export declare function makeXmlViewProps(): string;
export declare function getShapeInfo(shapeName: any): any;
