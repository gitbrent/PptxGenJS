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
		title:      'Slide master with corporate branding',
		isNumbered: true,
		margin:     [ 0.5, 0.25, 1.0, 0.25 ],
		bkgd:       'FFFFFF',
		images:     [
			{ x:9.3, y:4.9, cx:0.5, cy:0.5, src:'images/logo_square.png' },
			{ x:1.3, y:4.9, cx:0.5, cy:0.5, data:'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAABGRJREFUeNrsWE9IFGEU/3Zd3VVLJ/+UWsuuZKViOUmJGORWUHRqLYkoSIvwUkRCEATWXjqG6K3TrlDkIZo6dSgYu0SHajWi665gGFjMmrUlFl/zlvfJzDQzO7s7G23sg8c3M7sz3/u9/98jpEhFKhLQFZn9hSi4E1efzFGZg4VqAQ4AeHsvUnf15rGCBVHt64l2XXhC3esbwgiqMEF0X3pBYS1oEL1X3/8fIGq3H47J93xBg2jae14qVBAEAhpiou3EHSrfDhU0iJ1npqhngy/0L8nmyARE5+DDlAVmJ49HVpY/jsiXCQxyHis5cCc+4/B+RlNvgOP4fBbXmb9qif2jH2iVdw8EN8QGrILMIazkAQuZi8f/wTsifiPv7gkbCC5PNWUgbAxuPypBzEdPxqOGwkxYZomDt77anaGCuJdtGc/wg0oQrcEJyYILcKiEcBoBeWwsbWnwYmYmVYJI9VDmjSAIThVs1v2G7YiJEJ4TLAU2gIAVG0E9EjFw/fhtCYGEDKwg2qF9TvPREAphagnUoJ47RjV7CCaCRnMJaNhM5Q5llXUiCAdnBj3TAwgodNgAMl/ndKwqKmIgaACWZb1wLu6jErK8dmsMNAyMNYAz8F0BN48aaDGIICT83Sygs85Ieq4ShtzPfN3AdwlajqLguabFAILk0p2JteTTeWnu68Lb1MW6xl2kcmNbn8G70GKcQyCP5N7JX7P1QLYFalrmcY3bWQIAm2kFnF54cy918fPHEllNfk6YbBxBBYjNB6+TjtN3OTlGbmZphQgqJKPUOoam1wtuwSBAdcc16xo6BKjWstVojm0Ch+4Utup71EId8CvqBZdm8xCCz5VEk/j7IwNIZoL5yt3S6327KPD+mipJr0bkaQQkMWs6Tf44jn27YUU82VjHeT1uAvxg9w7uVGOdmGMLEEA2y1oJjItgugMNpxB+BgNJRZ3rK6Wn3e0qC11+HyNTC5/6IQNlolVXlVto7G8NuCpKycpiMvVQejk//X3+Sz8KrQc05EyD9Jymk1TR7PK3EVlY1bOJ9mZYBrUtw0CfU0SFiBgLSi3z/uGuwKajLaS2z0eaBtpSvO3avoC2I7CaRolC8+Pob3ogImPxhfjSz19rD94tJ4lSY/UcERYflQbvj5YEVp+VphjuD3U5REV8JZLxJZKcUyu6rL6CGGSujAsjaw1Y86Vq8loqPFHZ/+mwdxOtK3Opfu/a7pBkoamWb5wtoRpBUhmttNojbOjZIjUNtFNYdYRVBXE2IEIZdopDZ4846as7rjXhn9120R1eRzTLDGS5FhiBYMfLGLE+jmetuNDuc7A8nunIkreS2q0ewNkkIYbX+Zwo+BWnOVvPywEFEKqwzJANkwUOvyMqjqC67YvDJiCDuCaw+LGhFlw/x+czabTMhmK8RgFQtCaxMyX5AKDV2jEUIo5MFNO4OWzV/Zp6w35nwoOwj7EYJmwZLWbpuzyyz4JbxXHUGEkndJGKVKQi/Tv0W4ABAEiWpmHCFivDAAAAAElFTkSuQmCC' }
		],
		shapes:     [
			{ type:'text', text:'ACME - Confidential', x:0, y:5.17, cx:'100%', cy:0.3, align:'center', valign:'top', color:'7F7F7F', font_size:8, bold:true }
		]
	},
	TITLE_SLIDE: {
		title:      'Presentation Title Slide',
		isNumbered: false,
		bkgd:       { src:'images/title_bkgd.png' },
		images:     [ { src:'images/sample_logo.png', x:'7.4', y:'4.1', cx:'2', cy:'1' } ],
		shapes:     [
			{ type:'text', x:0.3, y:3.30, cx:5.5, cy:0.5, text:'Global IT Team', font_face:'Arial', color:'888888', font_size:20 },
			{ type:'line', x:0.3, y:3.85, cx:5.7, cy:0.0, line:'007AAA' },
			{ type:'rectangle', x:0, y:0, w:'100%', h:0.5, fill:'003b75' }
		]
	}
};
