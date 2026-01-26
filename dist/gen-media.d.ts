/**
 * PptxGenJS: Media Methods
 */
import { PresSlide, SlideLayout } from './core-interfaces';
/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export declare function encodeSlideMediaRels(layout: PresSlide | SlideLayout): Array<Promise<string>>;
/**
 * FIXME: TODO: currently unused
 * TODO: Should return a Promise
 */
