/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */
import { AlignH, AlignV, CHART_TYPE, ChartType, OutputType, SCHEME_COLOR_NAMES, SHAPE_TYPE, SchemeColor, ShapeType, WRITE_OUTPUT_TYPE } from './core-enums';
import { AddSlideProps, IPresentationProps, PresLayout, PresSlide, SectionProps, SlideLayout, SlideMasterProps, TableToSlidesProps, ThemeProps, WriteBaseProps, WriteFileProps, WriteProps } from './core-interfaces';
export default class PptxGenJS implements IPresentationProps {
    /**
     * Presentation layout name
     * Standard layouts:
     * - 'LAYOUT_4x3'   (10"    x 7.5")
     * - 'LAYOUT_16x9'  (10"    x 5.625")
     * - 'LAYOUT_16x10' (10"    x 6.25")
     * - 'LAYOUT_WIDE'  (13.33" x 7.5")
     * Custom layouts:
     * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
     * @type {string}
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     */
    private _layout;
    set layout(value: string);
    get layout(): string;
    /**
     * PptxGenJS Library Version
     */
    private readonly _version;
    get version(): string;
    /**
     * @type {string}
     */
    private _author;
    set author(value: string);
    get author(): string;
    /**
     * @type {string}
     */
    private _company;
    set company(value: string);
    get company(): string;
    /**
     * @type {string}
     * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
     */
    private _revision;
    set revision(value: string);
    get revision(): string;
    /**
     * @type {string}
     */
    private _subject;
    set subject(value: string);
    get subject(): string;
    /**
     * @type {ThemeProps}
     */
    private _theme;
    set theme(value: ThemeProps);
    get theme(): ThemeProps;
    /**
     * @type {string}
     */
    private _title;
    set title(value: string);
    get title(): string;
    /**
     * Whether Right-to-Left (RTL) mode is enabled
     * @type {boolean}
     */
    private _rtlMode;
    set rtlMode(value: boolean);
    get rtlMode(): boolean;
    /** master slide layout object */
    private readonly _masterSlide;
    get masterSlide(): PresSlide;
    /** this Presentation's Slide objects */
    private readonly _slides;
    get slides(): PresSlide[];
    /** this Presentation's sections */
    private readonly _sections;
    get sections(): SectionProps[];
    /** slide layout definition objects, used for generating slide layout files */
    private readonly _slideLayouts;
    get slideLayouts(): SlideLayout[];
    private LAYOUTS;
    private readonly _alignH;
    get AlignH(): typeof AlignH;
    private readonly _alignV;
    get AlignV(): typeof AlignV;
    private readonly _chartType;
    get ChartType(): typeof ChartType;
    private readonly _outputType;
    get OutputType(): typeof OutputType;
    private _presLayout;
    get presLayout(): PresLayout;
    private readonly _schemeColor;
    get SchemeColor(): typeof SchemeColor;
    private readonly _shapeType;
    get ShapeType(): typeof ShapeType;
    /**
     * @depricated use `ChartType`
     */
    private readonly _charts;
    get charts(): typeof CHART_TYPE;
    /**
     * @depricated use `SchemeColor`
     */
    private readonly _colors;
    get colors(): typeof SCHEME_COLOR_NAMES;
    /**
     * @depricated use `ShapeType`
     */
    private readonly _shapes;
    get shapes(): typeof SHAPE_TYPE;
    constructor();
    /**
     * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
     * @param {AddSlideProps} options - slide masterName and/or sectionTitle
     * @return {PresSlide} new Slide
     */
    private readonly addNewSlide;
    /**
     * Provides an API for `addTableDefinition` to get slide reference by number
     * @param {number} slideNum - slide number
     * @return {PresSlide} Slide
     * @since 3.0.0
     */
    private readonly getSlide;
    /**
     * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
     * @param {SlideNumberProps} slideNum - slide number config
     */
    private readonly setSlideNumber;
    /**
     * Create all chart and media rels for this Presentation
     * @param {PresSlide | SlideLayout} slide - slide with rels
     * @param {JSZip} zip - JSZip instance
     * @param {Promise<string>[]} chartPromises - promise array
     */
    private readonly createChartMediaRels;
    /**
     * Create and export the .pptx file
     * @param {string} exportName - output file type
     * @param {Blob} blobContent - Blob content
     * @return {Promise<string>} Promise with file name
     */
    private readonly writeFileToBrowser;
    /**
     * Create and export the .pptx file
     * @param {WRITE_OUTPUT_TYPE} outputType - output file type
     * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
     */
    private readonly exportPresentation;
    /**
     * Export the current Presentation to stream
     * @param {WriteBaseProps} props - output properties
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
     */
    stream(props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>;
    /**
     * Export the current Presentation as JSZip content with the selected type
     * @param {WriteProps} props output properties
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
     */
    write(props?: WriteProps | WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>;
    /**
     * Export the current Presentation.
     * Write the generated presentation to disk (Node) or trigger a download (browser).
     * @param {WriteFileProps} props - output file properties
     * @returns {Promise<string>} the presentation name
     */
    writeFile(props?: WriteFileProps | string): Promise<string>;
    /**
     * Add a new Section to Presentation
     * @param {ISectionProps} section - section properties
     * @example pptx.addSection({ title:'Charts' });
     */
    addSection(section: SectionProps): void;
    /**
     * Add a new Slide to Presentation
     * @param {AddSlideProps} options - slide options
     * @returns {PresSlide} the new Slide
     */
    addSlide(options?: AddSlideProps): PresSlide;
    /**
     * Create a custom Slide Layout in any size
     * @param {PresLayout} layout - layout properties
     * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
     */
    defineLayout(layout: PresLayout): void;
    /**
     * Create a new slide master [layout] for the Presentation
     * @param {SlideMasterProps} props - layout properties
     */
    defineSlideMaster(props: SlideMasterProps): void;
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * @param {string} eleId - table HTML element ID
     * @param {TableToSlidesProps} options - generation options
     */
    tableToSlides(eleId: string, options?: TableToSlidesProps): void;
}
