import { ONEPT } from './core-enums'
import { IGlowOptions } from './core-interfaces'
import { createColorElement, getMix } from './gen-utils'

// TODO: move createShadowElement here

/**
 * Creates `a:glow` element
 * @param {Object} opts glow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: IGlowOptions, defaults: IGlowOptions): string {
	var
		strXml = '',
		opts   = getMix(defaults, options),
		size           = opts['size'] * ONEPT,
		color           = opts['color'],
		opacity         = opts['opacity'] * 100000;

	strXml += '<a:glow rad="' + size + '">';
	strXml += createColorElement(color, '<a:alpha val="'+ opacity +'"/>');
	strXml += '</a:glow>';

	return strXml;
}