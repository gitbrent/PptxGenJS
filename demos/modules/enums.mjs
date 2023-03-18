/**
 * enums.mjs
 * global variables
 */

// LIBRARY
export const TESTMODE = typeof window !== "undefined" && window.location && window.location.href.toLowerCase().indexOf("http://localhost:8000/") > -1;
export const COMPRESS = true; // TEST: `compression` write prop

// CONST
export const CUST_NAME = "S.T.A.R. Laboratories";
export const USER_NAME = "Barry Allen";
export const ARRSTRBITES = [130];
export const CHARSPERLINE = 130; // "Open Sans", 13px, 900px-colW = ~19 words/line ~130 chars/line

// TABLES
export const TABLE_NAMES_F = ["Markiplier", "Jack", "Brian", "Paul", "Ev", "Ann", "Michelle", "Jenny", "Lara", "Kathryn"];
export const TABLE_NAMES_L = ["Johnson", "Septiceye", "Lapston", "Lewis", "Clark", "Griswold", "Hart", "Cube", "Malloy", "Capri"];
export const BASE_TABLE_OPTS = { x: 0.5, y: 0.13, colW: [9, 3.33] }; // LAYOUT_WIDE w=12.33

// STYLES
export const BKGD_LTGRAY = "F1F1F1";
export const COLOR_BLUE = "0088CC";
export const CODE_STYLE = { fill: { color: BKGD_LTGRAY }, margin: 6, fontSize: 10, color: '696969' };
export const TITLE_STYLE = { fill: { color: BKGD_LTGRAY }, margin: 4, fontSize: 18, fontFace: "Segoe UI", color: COLOR_BLUE, valign: "top", align: "center" };

// OPTIONS
export const BASE_TEXT_OPTS_L = { color: "9F9F9F", margin: 3, border: [null, null, { pt: "1", color: "CFCFCF" }, null] };
export const BASE_TEXT_OPTS_R = {
	text: "PptxGenJS",
	options: { color: "9F9F9F", margin: 3, border: [0, 0, { pt: "1", color: "CFCFCF" }, 0], align: "right" },
};
export const FOOTER_TEXT_OPTS = { x: 0.0, y: 7.16, w: "100%", h: 0.3, margin: 3, color: "9F9F9F", align: "center", fontSize: 10 };
export const BASE_CODE_OPTS = {
	color: "9F9F9F",
	margin: 3,
	border: { pt: "1", color: "CFCFCF" },
	fill: { color: "F1F1F1" },
	fontFace: "Courier",
	fontSize: 12,
};
export const BASE_OPTS_SUBTITLE = { x: 0.5, y: 0.7, w: 4, h: 0.3, fontSize: 18, fontFace: "Arial", color: "0088CC", fill: { color: "FFFFFF" } };
export const DEMO_TITLE_TEXT = { fontSize: 14, color: "0088CC", bold: true };
export const DEMO_TITLE_TEXTBK = { fontSize: 14, color: "0088CC", bold: true, breakLine: true };
export const DEMO_TITLE_OPTS = { fontSize: 13, color: "9F9F9F" };

// PATHS
export const IMAGE_PATHS = {
	peace4: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/peace4.png" },
	starlabsBkgd: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/starlabs_bkgd.jpg" },
	starlabsLogo: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/starlabs_logo.png" },
	wikimedia1: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/wiki-example.jpg" },
	wikimedia2: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/png-gradient-hex.png" },
	wikimedia_svg: { path: "https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@master/demos/common/images/lock-green.svg" },
	ccCopyRemix: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/cc_copyremix.gif" },
	ccLogo: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/cc_logo.jpg" },
	ccLicenseComp: { path: "/common/images/cc_license_comp.png" },
	ccDjGif: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/cc_dj.gif" },
	gifAnimTrippy: { path: "https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/demos/common/images/trippy.gif" },
	chicagoBean: {
		path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/chicago_bean_bohne.jpg?op=paramTest&ampersandTest",
	},
	sydneyBridge: {
		path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/sydney_harbour_bridge_night.jpg?op=paramTest&ampersandTest&fileType=.jpg",
	},
	tokyoSubway: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/tokyo-subway-route-map.jpg" },
	sample_aif: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.aif" },
	sample_avi: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.avi" },
	sample_m4v: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.m4v" },
	sample_mov: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.mov" },
	sample_mp4: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.mp4" },
	sample_mpg: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.mpg" },
	sample_mp3: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.mp3" },
	sample_wav: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/media/sample.wav" },
	big_earth_mp4: { path: "/common/media/earth-big.mp4" },
	UPPERCASE: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/UPPERCASE.PNG" },
	video_mp4_thumb: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/video-mp4-thumb.png" },
};

// LOREM IPSUM
export const LOREM_IPSUM_ENG =
	'Far far away, behind the word mountains, far from the countries Vokalia and Consonantia, there live the blind texts. Separated they live in Bookmarksgrove right at the coast of the Semantics, a large language ocean. A small river named Duden flows by their place and supplies it with the necessary regelialia. It is a paradisematic country, in which roasted parts of sentences fly into your mouth. Even the all-powerful Pointing has no control about the blind texts it is an almost unorthographic life One day however a small line of blind text by the name of Lorem Ipsum decided to leave for the far World of Grammar. The Big Oxmox advised her not to do so, because there were thousands of bad Commas, wild Question Marks and devious Semikoli, but the Little Blind Text didn’t listen. She packed her seven versalia, put her initial into the belt and made herself on the way. When she reached the first hills of the Italic Mountains, she had a last view back on the skyline of her hometown Bookmarksgrove, the headline of Alphabet Village and the subline of her own road, the Line Lane. Pityful a rethoric question ran over her cheek, then she continued her way. On her way she met a copy. The copy warned the Little Blind Text, that where it came from it would have been rewritten a thousand times and everything that was left from its origin would be the word "and" and the Little Blind Text should turn around and return to its own, safe country. But nothing the copy said could convince her and so it didn’t take long until a few insidious Copy Writers ambushed her, made her drunk with Longe and Parole and dragged her into their agency, where they abused her for their projects again and again. And if she hasn’t been rewritten, then they are still using her. Far far away, behind the word mountains, far from the countries Vokalia and Consonantia, there live the blind texts. Separated they live in Bookmarksgrove right at the coast of the Semantics, a large language ocean. A small river named Duden flows by their place and supplies it with the necessary regelialia. It is a paradisematic country, in which roasted parts of sentences fly into your mouth. Even the all-powerful Pointing has no control about the blind texts it is an almost unorthographic life One day however a small line of blind text by the name of Lorem Ipsum decided to leave for the far World of Grammar. The Big Oxmox advised her not to do so, because there were thousands of bad Commas, wild Question Marks and devious Semikoli, but the Little Blind Text didn’t listen. She packed her seven versalia, put her initial into the belt and made herself on the way. When she reached the first hills of the Italic Mountains, she had a last view back on the skyline of her hometown Bookmarksgrove, the headline of Alphabet Village and the subline of her own road, the Line Lane. Pityful a rethoric question ran over her cheek, then she continued her way. On her way she met a copy. The copy warned the Little Blind Text, that.';

export const LOREM_IPSUM =
	"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin condimentum dignissim velit vel luctus. Donec feugiat ipsum quis tempus blandit. Donec mattis mauris vel est dictum interdum. Pellentesque imperdiet nibh vitae porta ornare. Fusce non nisl lacus. Curabitur ut mattis dui. Ut pulvinar urna velit, vitae aliquam neque pulvinar eu. Fusce eget tellus eu lorem finibus mattis. Nunc blandit consequat arcu. Ut sed pharetra tortor, nec finibus ipsum. Pellentesque a est vitae ligula imperdiet rhoncus. Ut quis hendrerit tellus. Phasellus non malesuada mi. Suspendisse ullamcorper tristique odio fermentum elementum. Phasellus mattis mollis mauris, non mattis ligula dapibus quis. Quisque pretium metus massa. Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget. Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra... Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget. Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra.";
