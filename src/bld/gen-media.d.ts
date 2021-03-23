/**
 * PptxGenJS: Media Methods
 */
import { PresSlide, SlideLayout } from './core-interfaces';
/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export declare function encodeSlideMediaRels(layout: PresSlide | SlideLayout): Promise<string>[];
