import { inch2Emu } from '../utils/helpers.js';

export function moveTo(points){
    let aPoints = points.substr(1).split(','),
        moveTo = [`<a:moveTo>`,
            `<a:pt x="${inch2Emu(aPoints[0] / 96)}" y="${inch2Emu(aPoints[1] / 96)}"/>`,
            `</a:moveTo>`
        ];
    return moveTo.join('')
}

export function lnTo(points){
    let aPoints = points.substr(1).split(','),
        lineTo = [  `<a:lnTo>`,
            `<a:pt x="${ inch2Emu(aPoints[0] / 96)}" y="${ inch2Emu(aPoints[1] / 96)}"/>`,
            `</a:lnTo>`
        ];
    return lineTo.join('')
}

export function quadBezTo(points){
    let aPoints = points.substr(1).split(' '),
        quadBezTo = '<a:quadBezTo>';
    for (let i=0; i<aPoints.length; i++){
        let apts = aPoints[i].split(',');
        quadBezTo +=`<a:pt x="${inch2Emu(apts[0] / 96)}" y="${inch2Emu(apts[1] / 96)}"/>`;
    }

    quadBezTo += '</a:quadBezTo>';
    return quadBezTo;
}