"use strict";(self.webpackChunkpptxgenjs_gh_pages=self.webpackChunkpptxgenjs_gh_pages||[]).push([[4700],{3905:(e,t,a)=>{a.d(t,{Zo:()=>d,kt:()=>m});var n=a(7294);function r(e,t,a){return t in e?Object.defineProperty(e,t,{value:a,enumerable:!0,configurable:!0,writable:!0}):e[t]=a,e}function l(e,t){var a=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),a.push.apply(a,n)}return a}function i(e){for(var t=1;t<arguments.length;t++){var a=null!=arguments[t]?arguments[t]:{};t%2?l(Object(a),!0).forEach((function(t){r(e,t,a[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(a)):l(Object(a)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(a,t))}))}return e}function o(e,t){if(null==e)return{};var a,n,r=function(e,t){if(null==e)return{};var a,n,r={},l=Object.keys(e);for(n=0;n<l.length;n++)a=l[n],t.indexOf(a)>=0||(r[a]=e[a]);return r}(e,t);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(e);for(n=0;n<l.length;n++)a=l[n],t.indexOf(a)>=0||Object.prototype.propertyIsEnumerable.call(e,a)&&(r[a]=e[a])}return r}var s=n.createContext({}),p=function(e){var t=n.useContext(s),a=t;return e&&(a="function"==typeof e?e(t):i(i({},t),e)),a},d=function(e){var t=p(e.components);return n.createElement(s.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},c=n.forwardRef((function(e,t){var a=e.components,r=e.mdxType,l=e.originalType,s=e.parentName,d=o(e,["components","mdxType","originalType","parentName"]),c=p(a),m=r,y=c["".concat(s,".").concat(m)]||c[m]||u[m]||l;return a?n.createElement(y,i(i({ref:t},d),{},{components:a})):n.createElement(y,i({ref:t},d))}));function m(e,t){var a=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var l=a.length,i=new Array(l);i[0]=c;var o={};for(var s in t)hasOwnProperty.call(t,s)&&(o[s]=t[s]);o.originalType=e,o.mdxType="string"==typeof e?e:r,i[1]=o;for(var p=2;p<l;p++)i[p]=a[p];return n.createElement.apply(null,i)}return n.createElement.apply(null,a)}c.displayName="MDXCreateElement"},8114:(e,t,a)=>{a.r(t),a.d(t,{assets:()=>s,contentTitle:()=>i,default:()=>u,frontMatter:()=>l,metadata:()=>o,toc:()=>p});var n=a(7462),r=(a(7294),a(3905));const l={id:"usage-pres-options",title:"Presentation Options"},i=void 0,o={unversionedId:"usage-pres-options",id:"usage-pres-options",title:"Presentation Options",description:"Metadata",source:"@site/docs/usage-pres-options.md",sourceDirName:".",slug:"/usage-pres-options",permalink:"/PptxGenJS/docs/usage-pres-options",draft:!1,tags:[],version:"current",frontMatter:{id:"usage-pres-options",title:"Presentation Options"},sidebar:"docs",previous:{title:"Creating a Presentation",permalink:"/PptxGenJS/docs/usage-pres-create"},next:{title:"Adding a Slide",permalink:"/PptxGenJS/docs/usage-add-slide"}},s={},p=[{value:"Metadata",id:"metadata",level:2},{value:"Metadata Properties",id:"metadata-properties",level:3},{value:"Metadata Properties Examples",id:"metadata-properties-examples",level:3},{value:"Slide Layouts (Sizes)",id:"slide-layouts-sizes",level:2},{value:"Slide Layout Syntax",id:"slide-layout-syntax",level:3},{value:"Standard Slide Layouts",id:"standard-slide-layouts",level:3},{value:"Custom Slide Layouts",id:"custom-slide-layouts",level:3},{value:"Custom Slide Layout Example",id:"custom-slide-layout-example",level:3},{value:"Text Direction",id:"text-direction",level:2},{value:"Text Direction Options",id:"text-direction-options",level:3},{value:"Text Direction Examples",id:"text-direction-examples",level:3}],d={toc:p};function u(e){let{components:t,...a}=e;return(0,r.kt)("wrapper",(0,n.Z)({},d,a,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"metadata"},"Metadata"),(0,r.kt)("h3",{id:"metadata-properties"},"Metadata Properties"),(0,r.kt)("p",null,"There are several optional PowerPoint metadata properties that can be set."),(0,r.kt)("h3",{id:"metadata-properties-examples"},"Metadata Properties Examples"),(0,r.kt)("p",null,"PptxGenJS uses ES6-style getters/setters."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-javascript"},"pptx.author = 'Brent Ely';\npptx.company = 'S.T.A.R. Laboratories';\npptx.revision = '15';\npptx.subject = 'Annual Report';\npptx.title = 'PptxGenJS Sample Presentation';\n")),(0,r.kt)("h2",{id:"slide-layouts-sizes"},"Slide Layouts (Sizes)"),(0,r.kt)("p",null,"Layout Option applies to all the Slides in the current Presentation."),(0,r.kt)("h3",{id:"slide-layout-syntax"},"Slide Layout Syntax"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-javascript"},"pptx.layout = 'LAYOUT_NAME';\n")),(0,r.kt)("h3",{id:"standard-slide-layouts"},"Standard Slide Layouts"),(0,r.kt)("table",null,(0,r.kt)("thead",{parentName:"table"},(0,r.kt)("tr",{parentName:"thead"},(0,r.kt)("th",{parentName:"tr",align:"left"},"Layout Name"),(0,r.kt)("th",{parentName:"tr",align:"left"},"Default"),(0,r.kt)("th",{parentName:"tr",align:"left"},"Layout Slide Size"))),(0,r.kt)("tbody",{parentName:"table"},(0,r.kt)("tr",{parentName:"tbody"},(0,r.kt)("td",{parentName:"tr",align:"left"},(0,r.kt)("inlineCode",{parentName:"td"},"LAYOUT_16x9")),(0,r.kt)("td",{parentName:"tr",align:"left"},"Yes"),(0,r.kt)("td",{parentName:"tr",align:"left"},"10 x 5.625 inches")),(0,r.kt)("tr",{parentName:"tbody"},(0,r.kt)("td",{parentName:"tr",align:"left"},(0,r.kt)("inlineCode",{parentName:"td"},"LAYOUT_16x10")),(0,r.kt)("td",{parentName:"tr",align:"left"},"No"),(0,r.kt)("td",{parentName:"tr",align:"left"},"10 x 6.25 inches")),(0,r.kt)("tr",{parentName:"tbody"},(0,r.kt)("td",{parentName:"tr",align:"left"},(0,r.kt)("inlineCode",{parentName:"td"},"LAYOUT_4x3")),(0,r.kt)("td",{parentName:"tr",align:"left"},"No"),(0,r.kt)("td",{parentName:"tr",align:"left"},"10 x 7.5 inches")),(0,r.kt)("tr",{parentName:"tbody"},(0,r.kt)("td",{parentName:"tr",align:"left"},(0,r.kt)("inlineCode",{parentName:"td"},"LAYOUT_WIDE")),(0,r.kt)("td",{parentName:"tr",align:"left"},"No"),(0,r.kt)("td",{parentName:"tr",align:"left"},"13.3 x 7.5 inches")))),(0,r.kt)("h3",{id:"custom-slide-layouts"},"Custom Slide Layouts"),(0,r.kt)("p",null,"Custom, user-defined layouts are supported"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Use the ",(0,r.kt)("inlineCode",{parentName:"li"},"defineLayout()")," method to create a custom layout of any size"),(0,r.kt)("li",{parentName:"ul"},"Create as many layouts as needed, ex: create an 'A3' and 'A4' and set layouts as desired")),(0,r.kt)("h3",{id:"custom-slide-layout-example"},"Custom Slide Layout Example"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-javascript"},"// Define new layout for the Presentation\npptx.defineLayout({ name:'A3', width:16.5, height:11.7 });\n\n// Set presentation to use new layout\npptx.layout = 'A3'\n")),(0,r.kt)("h2",{id:"text-direction"},"Text Direction"),(0,r.kt)("h3",{id:"text-direction-options"},"Text Direction Options"),(0,r.kt)("p",null,"Right-to-Left (RTL) text is supported. Simply set the RTL mode at the Presentation-level."),(0,r.kt)("h3",{id:"text-direction-examples"},"Text Direction Examples"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-javascript"},"// Set right-to-left text mode\npptx.rtlMode = true;\n")),(0,r.kt)("p",null,"Notes:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"You may also need to set an RTL lang value such as ",(0,r.kt)("inlineCode",{parentName:"li"},"lang='he'")," as the default lang is 'EN-US'"),(0,r.kt)("li",{parentName:"ul"},"See ",(0,r.kt)("a",{parentName:"li",href:"https://github.com/gitbrent/PptxGenJS/issues/600"},"Issue#600")," for more")))}u.isMDXComponent=!0}}]);