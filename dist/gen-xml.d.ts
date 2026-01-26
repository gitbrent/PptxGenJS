/**
 * PptxGenJS: XML Generation
 */
import { IPresentationProps, ISlideObject, PresSlide, ShadowProps, SlideLayout, TableCell } from './core-interfaces';
/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {ISlideObject|TableCell} slideObj - slideObj or tableCell
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragragh-level formatting device
 * @template
 *    <p:txBody>
 *        <a:bodyPr wrap="square" rtlCol="0">
 *            <a:spAutoFit/>
 *        </a:bodyPr>
 *        <a:lstStyle/>
 *        <a:p>
 *            <a:pPr algn="ctr"/>
 *            <a:r>
 *                <a:rPr lang="en-US" dirty="0" err="1"/>
 *                <a:t>textbox text</a:t>
 *            </a:r>
 *            <a:endParaRPr lang="en-US" dirty="0"/>
 *        </a:p>
 *    </p:txBody>
 * @returns XML containing the param object's text and formatting
 */
export declare function genXmlTextBody(slideObj: ISlideObject | TableCell): string;
/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
export declare function genXmlPlaceholder(placeholderObj: ISlideObject): string;
/**
 * Generate XML ContentType
 * @param {PresSlide[]} slides - slides
 * @param {SlideLayout[]} slideLayouts - slide layouts
 * @param {PresSlide} masterSlide - master slide
 * @returns XML
 */
export declare function makeXmlContTypes(slides: PresSlide[], slideLayouts: SlideLayout[], masterSlide?: PresSlide): string;
/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export declare function makeXmlRootRels(): string;
/**
 * Creates `docProps/app.xml`
 * @param {PresSlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export declare function makeXmlApp(slides: PresSlide[], company: string): string;
/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} subject - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
export declare function makeXmlCore(title: string, subject: string, author: string, revision: string): string;
/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlide[]} slides - Presenation Slides
 * @returns XML
 */
export declare function makeXmlPresentationRels(slides: PresSlide[]): string;
/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export declare function makeXmlSlide(slide: PresSlide): string;
/**
 * Get text content of Notes from Slide
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
export declare function getNotesFromSlide(slide: PresSlide): string;
/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
export declare function makeXmlNotesMaster(): string;
/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export declare function makeXmlNotesSlide(slide: PresSlide): string;
/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
export declare function makeXmlLayout(layout: SlideLayout): string;
/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlide} slide - slide object that represents master slide layout
 * @param {SlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
export declare function makeXmlMaster(slide: PresSlide, layouts: SlideLayout[]): string;
/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export declare function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: SlideLayout[]): string;
/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlide[]} slides
 * @param {SlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export declare function makeXmlSlideRel(slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): string;
/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export declare function makeXmlNotesSlideRel(slideNumber: number): string;
/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlide} masterSlide - Slide object
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export declare function makeXmlMasterRel(masterSlide: PresSlide, slideLayouts: SlideLayout[]): string;
/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export declare function makeXmlNotesMasterRel(): string;
/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
export declare function makeXmlTheme(pres: IPresentationProps): string;
/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {IPresentationProps} pres - presentation
 * @return {string} XML
 */
export declare function makeXmlPresentation(pres: IPresentationProps): string;
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
/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} shadowProps - shadow options
 */
export declare function correctShadowOptions(shadowProps: ShadowProps): void;
