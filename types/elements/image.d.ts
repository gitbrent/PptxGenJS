import ElementInterface from './element-interface';
import Hyperlink, { HyperLinkOptions } from './hyperlink';
import Position, { PositionOptions } from './position';
export declare type ObjectFitOptions = 'none' | 'fill' | 'cover' | 'contain' | 'crop';
export declare type ColorBlend = {
    darkColor?: string;
    lightColor?: string;
};
declare type ImageFormat = {
    height: string | number;
    width: string | number;
};
export declare type ImageOptions = PositionOptions & {
    image?: string;
    rounding?: boolean;
    opacity?: number | string;
    placeholder?: string;
    colorBlend?: ColorBlend;
    objectFit?: ObjectFitOptions;
    imageFormat?: ImageFormat;
    data?: string;
    path?: string;
    hyperlink?: HyperLinkOptions;
};
export default class ImageElement implements ElementInterface {
    imgId: number;
    svgImgId: number;
    sourceH: any;
    sourceW: any;
    position: Position;
    image?: string;
    objectFit: ObjectFitOptions;
    imageFormat?: ImageFormat;
    rounding?: boolean;
    opacity?: number;
    colorBlend: ColorBlend;
    isSvg?: boolean;
    placeholder?: string;
    hyperlink?: Hyperlink;
    constructor(options: ImageOptions, relations: any);
    render(idx: any, presLayout: any, placeholder: any): string;
}
export {};
