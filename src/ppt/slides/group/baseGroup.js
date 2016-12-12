export default class BaseGroup {

    constructor( name ) {
        let id = new Date().getTime();
        this.wrapperGroupCoordinate = {id: id, name, x: 0, y: 0, cx: 0, cy: 0};
    }
}
