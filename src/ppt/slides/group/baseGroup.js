export default class BaseGroup {

    constructor( name ) {
        this.id = 2;
        this.wrapperGroupCoordinate = {id: this.id, name, x: 0, y: 0, cx: 0, cy: 0};
        this.groupStart = '';
        this.groupEnd = '';
    }
}
