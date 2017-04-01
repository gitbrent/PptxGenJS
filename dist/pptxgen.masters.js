/*\
|*|  :: pptxgen.masters.js ::
|*|
|*|  JavaScript framework that produces PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2017 Brent Ely -- https://github.com/gitbrent
|*|
|*|  Permission is hereby granted, free of charge, to any person obtaining a copy
|*|  of this software and associated documentation files (the "Software"), to deal
|*|  in the Software without restriction, including without limitation the rights
|*|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
|*|  copies of the Software, and to permit persons to whom the Software is
|*|  furnished to do so, subject to the following conditions:
|*|
|*|  The above copyright notice and this permission notice shall be included in all
|*|  copies or substantial portions of the Software.
|*|
|*|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
|*|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
|*|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
|*|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
|*|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
|*|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
|*|  SOFTWARE.
\*/

/*
 INSTRUCTIONS:
  - Use the 3 pre-defined Master Slides below as examples to get started
  - Add your own objects with your selected name to be able to use them in method calls
    EX: var slide1 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE );

 NOTES:
  - DO NOT change the 'gObjPptxMasters' variable name or the library wont be able to see your slides
  - This file can be named anything as long as its included in your html file before PptxGenJS.js
  EX:
    <script type="text/javascript" src="../dist/pptxgen.masters.js"></script>
    <script type="text/javascript" src="../dist/pptxgen.js"></script>
 */

var gObjPptxMasters = {
	TITLE_SLIDE: {
		title:  'Corporate Presentation Title Slide',
		bkgd:   { src:'images/title_bkgd.png' },
		images: [
			{ src:'images/sample_logo.png', x:10.3, y:5.6, w:2.0, h:1.0 }
		],
		shapes: [
			{ type:'rectangle', x:0.0, y:0.00, w:'100%', h:0.5, fill:'003b75' },
			{ type:'text',      x:0.3, y:3.30, w:5.5,    h:0.5, text:'Global IT Team', font_face:'Arial', color:'888888', font_size:20 },
			{ type:'line',      x:0.3, y:3.85, w:5.7,    h:0.0, line:'007AAA' }
		]
	},
	MASTER_SLIDE: {
		title:   'Corporate Branded Slide',
		bkgd:    'FFFFFF',
		margin:  [ 0.5, 0.25, 1.0, 0.25 ],
		images:  [
			{ x:12.4, y:6.6, w:0.5, h:0.5, data:'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABkAAAAZCAYAAADE6YVjAAAFf0lEQVR42r2Ve0xTZxjGD2JEE2UGBZFroZRSoC1QLkXud7Bys3IRtqE4zKYbENnMGo1zXoYDdZBqBVQUdUBAYGC5FsbELWyI8YITUVS0Igo9p2yJczNbn33A4lwW2Qxub/LkO3985/2d533f7zsUNZMApUf9Z9EUZUCVR/pSFXHHqLIVVq82eRWlP7s6fBlVLVFSVbFP9aviMbsybvMry8+qC1xoWxeZa1Ir+XFBrQQGNbGYVyPF3Bpp8SsBBJ8JNvdvjFCKlFE6XkMUbOolsKiPg0VDMubWxnTPGLCmMdA0uSWsO6Y5HKGNEfBVRsJTKYF74yqImlOxqFYyNDNAZ+DczI6gr9arQpGhDMHOkgAo8v1QlBuAg3mhyFFEQVAVcWlGkO1d/vKtHQG62v3LoM7wAJ3igvGkCYkwnuCGsXgBHsQ4/syEcvLV3hZGLw3Y962X+IDK56cb2SIwJDm92h2MVABtDA/aFQ7QRjtCKxWRVQhtGE/HBHBGtF7WQS91yE52etb15/jgUlgEVM4rcMZaglYHCS4EheNhDEkeyYU2fALmBiZCAMafC1rM+UHjZrsS1L84pOW9fIeePX70hbB49PpJccUnEX3iJFwUJeAbQTxUrsuhjvKANpgLJsgBTBQB+TuC8bQH7cqhR/mW/v8I6aoXZ/SvidNdX5mIW3GrcScmFbclqbgVmYKBsGT0+EtxLjASmmCS2NeegJzBBAjBeBCokAPGid0zvHTp4mkh3+8Nk99JT8LQ2mTcezMF6tQU3CRSpqfq2t54/dr5uESVKlgycD/SS8d4E4iYJA90BePOA8PngHbgYIxtk0NNV7bbeVFNw1lJGH4vGcMbk9GRmQxpbhoExRuvW8rTzSb2tPj5Lb0WKr7LeJEvdyfyIxARgTgRiD0HGlv2TbXhNBM3sie66dGWBDySJaD3w0R4F6bBpeJ9CBSy6uf3PfDll08CXIm8SblcHSdd0Gw70Cy2TmPGin3B7PYJa/Zl92g+iYdm10q8m7caws83QdSwA4Itsn47u7XGE9vs7GTG19xchiYBAiIvMuZCJ+KCNN+GQCztoDGzLf3rpB0cnU8V3SjQk199kiE/ohsvjMbD/bHgK9bDrWYbPJv2wk2eq7O3e2fQkfdBAd9py5BS6KejCYCeKJEHgfCdiQvybEUgFkTm7L6bFGUwBTg2wppz4u53C4/fgmnpdfiWnMW9w1LcL5KAX5IFj/pceLXI4VmjAC8wCzyHTeA7ybDVOWkKwCVyEYJ25hMXnCkAkcac/QttZGRI/hHDVrMrR/oNTt7DvKODmFfUj/nyyzh0ZCs0p8LgWpIDj7q98Gouhk9rGcSKQ+B5ZILHzYbYMRs3HEiJOCQpj/TDmbhhcZ9BJvTIxJJNURVall7lw/ZZJ9S/6R0ZhJ6iH/oFl2G9rxN95YnYcOxtiKo/g3dTKfxbKxHUVgf/0goIEz4GzysH6S6pGOYSAIs4cCRubHmkH3+6GTW2dp0q1ynakDqu3kaVDI5SBKJHIAZ5vViWX4/qsnUIqcyDb0MZglpOI7StEeFtHQhv/RqhNecQcqILO7cX69R2/Kc0y4G4In2xsn8G0ZrYCJ+7r6BHFQ0ICeQwVXBlbE5+L0x3d8N7ZxOyD+yHpKoMIU0NJLkKEW3nEKHqRVT7VUQ39D+OOz2gOO8StHzMzPYsAf06CfsDNP7aEpu/j3BClT5VeJEzK683y3h3dzN3e9ewn0wFqawBq3Y1QFrYjARFx9PkA51X0j7tlKft+FKUMPEOCZoyMqSXsJI0lpwWjRXnMVmfjC1evGD6S+yjqjnmsvZFnpva7SMzm30TN3wRtOatGvG6lFPWm2OOLnjRtTExtlpTFos2sQkn50Sf+j/id3j8G3eDGa7zAAAAAElFTkSuQmCC' },
			{ x:12.3, y:0.3, w:0.7, h:0.7, data:'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' }
		],
		shapes:  [
			{ type:'text', text:'S.T.A.R. Laboratories - Confidential', x:0, y:6.8, w:'100%', h:0.3, align:'center', valign:'top', color:'7F7F7F', font_size:9, bold:true }
		],
		slideNumber: { x:0.3, y:'90%', color:'0088CC', fontFace:'Courier' }
	},
	THANKS_SLIDE: {
		title:   'Thanks Slide to be added as final Slide',
		bkgd:    '36ABFF',
		shapes:  [
			{ type:'rectangle', x:0.0, y:3.4, w:'100%', h:2.0, fill:'ffffff' },
			{ type:'text', text:'Thank You!', x:0.0, y:0.9, w:'100%', h:1, font_face:'Arial', color:'ffffff', font_size:60, align:'c' }
		],
		images:  [
			{ src:'images/starlabs_logo.png', x:4.6, y:3.5, w:4, h:1.8 }
		]
	}
};

if ( typeof module !== 'undefined' && module.exports ) {
	module.exports = gObjPptxMasters;
}
