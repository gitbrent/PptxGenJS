/*\
|*|  :: pptxgen.masters.js ::
|*|
|*|  A complete JavaScript PowerPoint presentation creator framework for client browsers.
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2016 Brent Ely -- https://github.com/gitbrent
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

var gObjPptxMasters = {
	MASTER_SLIDE: {
		title:      'Basic corp slide master',
		isNumbered: true,
		margin:     [ 0.5, 0.25, 1.0, 0.25 ],
		bkgd:       'FFFFFF',
		images:     [ { src:'images/logo_square.png', x:9.3, y:4.9, cx:0.5, cy:0.5 } ],
		shapes:     [
			{ type:'text', text:'ACME - Confidential', x:0, y:5.17, cx:'100%', cy:0.3, align:'center', valign:'top', color:'7F7F7F', font_size:8, bold:true }
		]
	},
	TITLE_SLIDE: {
		title:      'I am the Title Slide',
		isNumbered: false,
		bkgd:       { src:'images/title_bkgd.png' },
		images:     [ { src:'images/sample_logo.png', x:'7.4', y:'4.1', cx:'2', cy:'1' } ],
		shapes:     [
			{ type:'text', x:0.3, y:3.3, cx:5.5, cy:0.5, text:'Global IT Team', font_face:'Arial', color:'888888', font_size:20 },
			{ type:'line', x:0.3, y:3.85, cx:5.7, cy:0.0, line:'007AAA' }
		]
	}
};
