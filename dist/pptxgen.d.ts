// Type definitions for pptxgenjs 2.3.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 2.3

declare namespace PptxGenJS {
  interface ENUMS {
    chartTypes: "AREA" | "BAR" | "BUBBLE" | "DOUGHNUT" | "LINE" | "PIE" | "RADAR" | "SCATTER",
    jsZipOutputTypes: "arraybuffer" | "base64" | "binarystring" | "blob" | "nodebuffer" | "uint8array",
    layoutNames: "LAYOUT_4x3" | "LAYOUT_16x9" | "LAYOUT_16x10" | "LAYOUT_WIDE" | "LAYOUT_USER",
  }

  interface ImageOptions {
    x: number;
    y: number;
    w: number;
    h: number;

    base64data?: string;
    urlPath?: string;

    hyperlink?: string;
    sizing?: "cover" | "contain" | "crop";
  }
  interface MediaOptions {
    x: number;
    y: number;
    w: number;
    h: number;

    base64data?: string;
    urlPath?: string;

    onlineVideoLink?: string;
    type?: "audio" | "online" | "video";
  }

  const version: string;

  // Presentation Props
  function getLayout(): string;
  function setBrowser(isBrowser: boolean): void;
  function setLayout(layoutName: ENUMS["layoutNames"]): void;
  function setRTL(isRTL: boolean): void;

  // Presentation Metadata
  function setAuthor(author: string): void;
  function setCompany(company: string): void;
  function setRevision(revision: string): void;
  function setSubject(subject: string): void;
  function setTitle(title: string): void;

  class slide {
    // Slide Number methods
    getPageNumber(): string;
    slideNumber(): Object;
    slideNumber(options: Object): void;

    // Core Object API Methods
    addChart(type: ENUMS["chartTypes"], data: string, options?: Object): slide;
    addImage(options: ImageOptions): slide;
    addMedia(options: MediaOptions): slide;
    addNotes(noteText: string): slide;
    addShape(shapeName: string, options: Object): slide;
    addTable(tableData: Array<any>, options: Object): slide;
    addText(textString: string, options: Object): slide;
  }

  // Add a new Slide
  function addNewSlide(masterLayoutName?: string): slide;

  // Export
  function save(exportFileName: string, callbackFunction?: Function, zipOutputType?:ENUMS["jsZipOutputTypes"]): void;
}
