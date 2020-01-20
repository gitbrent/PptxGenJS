import ElementInterface from './element-interface';
import Relations from '../relations';
export default class MediaElement implements ElementInterface {
    videoId: any;
    mediaId: any;
    previewId: any;
    mediaType: any;
    media: any;
    position: any;
    constructor(options: any, relations: Relations);
    render(idx: any, presLayout: any): string;
}
