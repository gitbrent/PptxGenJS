/**
 * PptxGenJS: Media Methods
 */
import Slide from './slide';
import { Master } from './slideLayouts';
/**
 * Encode Image/Audio/Video into base64
 * @param {Slide | Master} layout - slide layout
 * @return {Promise} promise of generating the rels
 */
export declare function encodeSlideMediaRels(layout: Slide | Master): Promise<string>[];
