import ParagraphProperties, { ParagraphPropertiesOptions } from './paragraph-properties';
import RunProperties, { RunPropertiesOptions } from './run-properties';
import Relations from '../relations';
interface FragmentConfig {
    text: string;
    options?: ParagraphPropertiesOptions & RunPropertiesOptions;
}
export declare type FragmentOptions = string | number | FragmentConfig[];
export declare const buildFragments: (inputText: FragmentOptions, inputBreakLine: boolean, relations: Relations) => any;
export default class TextFragment {
    text: string;
    paragraphConfig: ParagraphProperties;
    runConfig: RunProperties;
    constructor(text: string, paragraphConfig: ParagraphProperties, runConfig: RunProperties);
    render(presLayout: any): string;
}
export {};
