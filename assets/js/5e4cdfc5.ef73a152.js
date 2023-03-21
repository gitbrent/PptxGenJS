"use strict";(self.webpackChunkpptxgenjs_gh_pages=self.webpackChunkpptxgenjs_gh_pages||[]).push([[9846],{4035:(e,t,n)=>{n.r(t),n.d(t,{default:()=>r});var o=n(7294),l=n(3285),a=n(9943);const r=()=>{const e=()=>o.createElement("section",{className:"bgTheme p-4"},o.createElement("h3",{className:"mb-3"},"About HTML-to-PPTX"),o.createElement("p",null,"The ",o.createElement("code",null,"tableToSlides")," method generates a presentation from an HTML table element id."),o.createElement("ul",null,o.createElement("li",null,"Many options are available including repeating header, start location on subsequent slides, character and line weight"),o.createElement("li",null,"Additional slides are automatically created as needed (auto-paging)"),o.createElement("li",null,"The table's style (CSS) is copied into the PowerPoint table")),o.createElement("div",{className:"d-none d-md-flex row align-items-center justify-content-center my-3"},o.createElement("div",{className:"col-auto"},o.createElement("img",{className:"d-none d-md-none d-lg-block border border-light",alt:"input: html table",src:"/PptxGenJS/img/ex-html-to-powerpoint-1.png",height:"400"}),o.createElement("img",{className:"d-none d-md-block d-lg-none border border-light",alt:"input: html table",src:"/PptxGenJS/img/ex-html-to-powerpoint-1.png",height:"300"}),o.createElement("img",{className:"d-block d-md-none d-lg-none border border-light",alt:"input: html table",src:"/PptxGenJS/img/ex-html-to-powerpoint-1.png",height:"200"})),o.createElement("div",{className:"col-auto"},o.createElement("h1",{className:"mb-0"},"\u2192")),o.createElement("div",{className:"col-auto"},o.createElement("img",{className:"d-none d-md-none d-lg-block border border-light",alt:"output: powerpoint slides",src:"/PptxGenJS/img/ex-html-to-powerpoint-2.png",height:"400"}),o.createElement("img",{className:"d-none d-md-block d-lg-none border border-light",alt:"output: powerpoint slides",src:"/PptxGenJS/img/ex-html-to-powerpoint-2.png",height:"300"}),o.createElement("img",{className:"d-block d-md-none d-lg-none border border-light",alt:"output: powerpoint slides",src:"/PptxGenJS/img/ex-html-to-powerpoint-2.png",height:"200"})))),t=()=>o.createElement("div",{className:"card useTheme h-100"},o.createElement("div",{className:"card-body"},o.createElement("h3",{className:"mb-3"},"Sample Code"),o.createElement("p",null,"Reproduce a table in as little as 3 lines of code."),o.createElement(a.Z,{id:"494850b6775dd5c8ce314672a1846208"})),o.createElement("div",{className:"card-footer text-center"},o.createElement("button",{type:"button","aria-label":"documentation",className:"btn btn-outline-primary px-5",onClick:()=>window.location.href="/PptxGenJS/docs/html-to-powerpoint"},"HTML to PowerPoint Docs"))),n=()=>o.createElement("div",{className:"card useTheme h-100"},o.createElement("div",{className:"card-body"},o.createElement("h3",{className:"mb-3"},"Live Demo"),o.createElement("p",null,"Try the html-to-pptx feature out for yourself."),o.createElement("div",{className:"text-center"},o.createElement("img",{alt:"HTML Table",src:"/PptxGenJS/img/ex-html-to-powerpoint-3.png",className:"border border-light"}))),o.createElement("div",{className:"card-footer text-center"},o.createElement("button",{type:"button","aria-label":"demo",className:"btn btn-outline-primary px-5",onClick:()=>window.location.href="/PptxGenJS/demo/browser/index.html#html2pptx"},"HTML to PowerPoint Demo")));return o.createElement(l.Z,{title:"HTML-to-PowerPoint"},o.createElement("div",{className:"container my-4"},o.createElement("h1",{className:"mb-4"},"HTML to PowerPoint"),o.createElement("div",{className:"row g-5"},o.createElement("div",{className:"col-12"},o.createElement(e,null)),o.createElement("div",{className:"col"},o.createElement(t,null)),o.createElement("div",{className:"col"},o.createElement(n,null)))))}},9943:(e,t,n)=>{n.d(t,{Z:()=>a});var o=n(7294);function l(e,t){return l=Object.setPrototypeOf||function(e,t){return e.__proto__=t,e},l(e,t)}const a=function(e){var t,n;function a(){return e.apply(this,arguments)||this}n=e,(t=a).prototype=Object.create(n.prototype),t.prototype.constructor=t,l(t,n);var r=a.prototype;return r.componentDidMount=function(){this._updateIframeContent()},r.componentDidUpdate=function(e,t){this._updateIframeContent()},r._defineUrl=function(){var e=this.props,t=e.id,n=e.file;return"https://gist.github.com/"+t+".js"+(n?"?file="+n:"")},r._updateIframeContent=function(){var e=this.props,t=e.id,n=e.file,o=this.iframeNode,l=o.document;o.contentDocument?l=o.contentDocument:o.contentWindow&&(l=o.contentWindow.document);var a='<html><head><base target="_parent"><style>*{font-size:12px;}</style></head><body '+("onload=\"parent.document.getElementById('"+(n?"gist-"+t+"-"+n:"gist-"+t)+"').style.height=document.body.scrollHeight + 'px'\"")+">"+('<script type="text/javascript" src="'+this._defineUrl()+'"><\/script>')+"</body></html>";l.open(),l.writeln(a),l.close()},r.render=function(){var e=this,t=this.props,n=t.id,l=t.file;return o.createElement("iframe",{ref:function(t){e.iframeNode=t},width:"100%",frameBorder:0,id:l?"gist-"+n+"-"+l:"gist-"+n})},a}(o.PureComponent)}}]);