import BaseGroup from './baseGroup';
import { inch2Emu } from '../utils/helpers';
import Slide from '../slide.js';

export default class SlideGroup extends BaseGroup {
    constructor() {
        super();
        this.allShapes = Slide.gObjPptx.slides;
    }

    set groupStart(sGroupStart){
      this._groupStart = sGroupStart;
    }

    get groupStart(){
      return this._groupStart;
    }

    set groupEnd(sGroupEnd){
      this._groupEnd = sGroupEnd;
    }

    get groupEnd(){
      return this._groupEnd;
    }

    generateGroup() {
      let nLength = this.allShapes.length, nIndex;

      for (nIndex=0; nIndex < nLength; nIndex++){
          let oShapes = this.allShapes[nIndex].data,
              groupIndex=0;

          for(let j=0; j< oShapes.length; j++){
              let oShape = oShapes[j].options;
              Object.keys(this.wrapperGroupCoordinate).map((i) => {
                  let nValue;
                  if (i === 'name'){
                      this.wrapperGroupCoordinate[i] = this.allShapes[nIndex].name;
                  }
                  // get the x,y of the first shape
                  else if (i === 'x' || i === 'y'){
                      if (groupIndex < 2){
                          this.wrapperGroupCoordinate[i] =  oShape[i];
                          groupIndex++;
                      } else {
                          // if an unordred shapes
                          this.wrapperGroupCoordinate[i] =  Math.min(this.wrapperGroupCoordinate[i] , oShape[i])
                      }
                  }
                  else if (i === 'cx' && (oShape[i] || oShape['w'])){
                      nValue =  oShape[i]? oShape[i]: oShape['w'];

                      if (groupIndex === 2){
                          this.wrapperGroupCoordinate[i] = nValue;
                          groupIndex++;
                      }

                      if (oShape['x'] > (this.wrapperGroupCoordinate[i] + this.wrapperGroupCoordinate['x']) ){
                          this.wrapperGroupCoordinate[i] = oShape['x'] + nValue;
                      }
                  }
                  else if(i === 'cy' && (oShape[i] || oShape['h'])){
                      nValue = oShape[i]? oShape[i]: oShape['h'];

                      if (groupIndex === 3){
                          this.wrapperGroupCoordinate[i] = nValue;
                          groupIndex++;
                      }
                      if ( oShape['y'] > ( this.wrapperGroupCoordinate[i] + this.wrapperGroupCoordinate['y']) ){
                          this.wrapperGroupCoordinate[i] = oShape['y'] + nValue;
                      }
                  }
              })
          }
      }
      let sStart = [
          `<p:grpSp>`,
          ` <p:nvGrpSpPr>`,
          ` <p:cNvPr id="${this.wrapperGroupCoordinate.id}" name="${this.wrapperGroupCoordinate.name}"/>`,
          `<p:cNvGrpSpPr/>`,
          `<p:nvPr/>`,
          `</p:nvGrpSpPr>`,
          `<p:grpSpPr>`,
          `<a:xfrm>`,
          `<a:off x="${inch2Emu(this.wrapperGroupCoordinate.x)}"  y="${inch2Emu(this.wrapperGroupCoordinate.y)}"/>`,
          `<a:ext cx="${inch2Emu(this.wrapperGroupCoordinate.cx)}" cy="${inch2Emu(this.wrapperGroupCoordinate.cy)}"/>`,
          `<a:chOff x="${inch2Emu(this.wrapperGroupCoordinate.x)}"  y="${inch2Emu(this.wrapperGroupCoordinate.y)}" />`,
          `<a:chExt cx="${inch2Emu(this.wrapperGroupCoordinate.cx)}" cy="${inch2Emu(this.wrapperGroupCoordinate.cy)}"/>`,
          `</a:xfrm>`,
          `</p:grpSpPr>`
      ];
        this.groupStart = sStart.join('')
        this.groupEnd = '</p:grpSp>';
        this.id++;
        return this;
    }

}
