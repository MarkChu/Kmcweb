(function(e){function n(n){for(var r,o,c=n[0],i=n[1],l=n[2],f=0,s=[];f<c.length;f++)o=c[f],Object.prototype.hasOwnProperty.call(a,o)&&a[o]&&s.push(a[o][0]),a[o]=0;for(r in i)Object.prototype.hasOwnProperty.call(i,r)&&(e[r]=i[r]);p&&p(n);while(s.length)s.shift()();return u.push.apply(u,l||[]),t()}function t(){for(var e,n=0;n<u.length;n++){for(var t=u[n],r=!0,o=1;o<t.length;o++){var c=t[o];0!==a[c]&&(r=!1)}r&&(u.splice(n--,1),e=i(i.s=t[0]))}return e}var r={},o={3:0},a={3:0},u=[];function c(e){return i.p+"js/"+({}[e]||e)+"."+{1:"ad667aa8",2:"f9fe835b",4:"6cd5a54a",5:"930adfa2",6:"a907e48c",7:"aadef51e",8:"0e4c9bf2",9:"1e8f24a0",10:"69c1fe3f",11:"7466bdf3",12:"89adcbf9",13:"a6d4660a"}[e]+".js"}function i(n){if(r[n])return r[n].exports;var t=r[n]={i:n,l:!1,exports:{}};return e[n].call(t.exports,t,t.exports,i),t.l=!0,t.exports}i.e=function(e){var n=[],t={1:1};o[e]?n.push(o[e]):0!==o[e]&&t[e]&&n.push(o[e]=new Promise((function(n,t){for(var r="css/"+({}[e]||e)+"."+{1:"b5a3905f",2:"31d6cfe0",4:"31d6cfe0",5:"31d6cfe0",6:"31d6cfe0",7:"31d6cfe0",8:"31d6cfe0",9:"31d6cfe0",10:"31d6cfe0",11:"31d6cfe0",12:"31d6cfe0",13:"31d6cfe0"}[e]+".css",a=i.p+r,u=document.getElementsByTagName("link"),c=0;c<u.length;c++){var l=u[c],f=l.getAttribute("data-href")||l.getAttribute("href");if("stylesheet"===l.rel&&(f===r||f===a))return n()}var s=document.getElementsByTagName("style");for(c=0;c<s.length;c++){l=s[c],f=l.getAttribute("data-href");if(f===r||f===a)return n()}var p=document.createElement("link");p.rel="stylesheet",p.type="text/css",p.onload=n,p.onerror=function(n){var r=n&&n.target&&n.target.src||a,u=new Error("Loading CSS chunk "+e+" failed.\n("+r+")");u.code="CSS_CHUNK_LOAD_FAILED",u.request=r,delete o[e],p.parentNode.removeChild(p),t(u)},p.href=a;var d=document.getElementsByTagName("head")[0];d.appendChild(p)})).then((function(){o[e]=0})));var r=a[e];if(0!==r)if(r)n.push(r[2]);else{var u=new Promise((function(n,t){r=a[e]=[n,t]}));n.push(r[2]=u);var l,f=document.createElement("script");f.charset="utf-8",f.timeout=120,i.nc&&f.setAttribute("nonce",i.nc),f.src=c(e);var s=new Error;l=function(n){f.onerror=f.onload=null,clearTimeout(p);var t=a[e];if(0!==t){if(t){var r=n&&("load"===n.type?"missing":n.type),o=n&&n.target&&n.target.src;s.message="Loading chunk "+e+" failed.\n("+r+": "+o+")",s.name="ChunkLoadError",s.type=r,s.request=o,t[1](s)}a[e]=void 0}};var p=setTimeout((function(){l({type:"timeout",target:f})}),12e4);f.onerror=f.onload=l,document.head.appendChild(f)}return Promise.all(n)},i.m=e,i.c=r,i.d=function(e,n,t){i.o(e,n)||Object.defineProperty(e,n,{enumerable:!0,get:t})},i.r=function(e){"undefined"!==typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},i.t=function(e,n){if(1&n&&(e=i(e)),8&n)return e;if(4&n&&"object"===typeof e&&e&&e.__esModule)return e;var t=Object.create(null);if(i.r(t),Object.defineProperty(t,"default",{enumerable:!0,value:e}),2&n&&"string"!=typeof e)for(var r in e)i.d(t,r,function(n){return e[n]}.bind(null,r));return t},i.n=function(e){var n=e&&e.__esModule?function(){return e["default"]}:function(){return e};return i.d(n,"a",n),n},i.o=function(e,n){return Object.prototype.hasOwnProperty.call(e,n)},i.p="",i.oe=function(e){throw console.error(e),e};var l=window["webpackJsonp"]=window["webpackJsonp"]||[],f=l.push.bind(l);l.push=n,l=l.slice();for(var s=0;s<l.length;s++)n(l[s]);var p=f;u.push([0,0]),t()})({0:function(e,n,t){e.exports=t("2f39")},"2f39":function(e,n,t){"use strict";t.r(n);var r={};t.r(r),t.d(r,"SET_Search",(function(){return P})),t.d(r,"SET_SearchRows",(function(){return x}));var o={};t.r(o),t.d(o,"setSearch",(function(){return S})),t.d(o,"setSearchRows",(function(){return k}));var a=t("c973"),u=t.n(a),c=(t("96cf"),t("5319"),t("ac1f"),t("5c7d"),t("573e"),t("7d6e"),t("e54f"),t("985d"),t("31cd"),t("2b0e")),i=t("1f91"),l=t("42d2"),f=t("b05d");c["a"].use(f["a"],{config:{},lang:i["a"],iconSet:l["a"]});var s=function(){var e=this,n=e.$createElement,t=e._self._c||n;return t("div",{attrs:{id:"q-app"}},[t("router-view")],1)},p=[],d={name:"App"},h=d,m=t("2877"),b=Object(m["a"])(h,s,p,!1,null,null,null),v=b.exports,w=t("2f62"),g=function(){return{Search:null,SearchRows:[]}},y=t("c091"),P=function(e,n){e.Search=n},x=function(e,n){e.SearchRows=n},S=function(e,n){var t=e.commit;t("SET_Search",n)},k=function(e,n){var t=e.commit;t("SET_SearchRows",n)},_={namespaced:!0,state:g,getters:y,mutations:r,actions:o};c["a"].use(w["a"]);var E=function(){var e=new w["a"].Store({modules:{law:_},strict:!1});return e},O=t("8c4f"),j=(t("e6cf"),t("d3b7"),t("3ca3"),t("e260"),t("ddb0"),[{path:"/",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(2)]).then(t.bind(null,"4648"))}}]},{path:"/acc1",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(2)]).then(t.bind(null,"4648"))}}]},{path:"/acc2",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(10)]).then(t.bind(null,"a8b7"))}}]},{path:"/acc3",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(11)]).then(t.bind(null,"174e"))}}]},{path:"/motion",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(12)]).then(t.bind(null,"6355"))}}]},{path:"/weblaw1",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(4)]).then(t.bind(null,"c5f8"))}}]},{path:"/weblaw2",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(5)]).then(t.bind(null,"5500"))}}]},{path:"/weblaw3",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(6)]).then(t.bind(null,"b034"))}}]},{path:"/weblaw4",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(7)]).then(t.bind(null,"a0bf"))}}]},{path:"/weblaw",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(8)]).then(t.bind(null,"ed04"))}}]},{path:"/weblawcontent",component:function(){return Promise.all([t.e(0),t.e(1)]).then(t.bind(null,"713b"))},children:[{path:"",component:function(){return Promise.all([t.e(0),t.e(13)]).then(t.bind(null,"3d06"))}}]},{path:"*",component:function(){return Promise.all([t.e(0),t.e(9)]).then(t.bind(null,"e51e"))}}]),T=j;c["a"].use(O["a"]);var R=function(){var e=new O["a"]({scrollBehavior:function(){return{x:0,y:0}},routes:T,mode:"hash",base:""});return e},C=function(){return A.apply(this,arguments)};function A(){return A=u()(regeneratorRuntime.mark((function e(){var n,t,r;return regeneratorRuntime.wrap((function(e){while(1)switch(e.prev=e.next){case 0:if("function"!==typeof E){e.next=6;break}return e.next=3,E({Vue:c["a"]});case 3:e.t0=e.sent,e.next=7;break;case 6:e.t0=E;case 7:if(n=e.t0,"function"!==typeof R){e.next=14;break}return e.next=11,R({Vue:c["a"],store:n});case 11:e.t1=e.sent,e.next=15;break;case 14:e.t1=R;case 15:return t=e.t1,n.$router=t,r={router:t,store:n,render:function(e){return e(v)}},r.el="#q-app",e.abrupt("return",{app:r,store:n,router:t});case 20:case"end":return e.stop()}}),e)}))),A.apply(this,arguments)}var L=t("b50c");c["a"].prototype.$apiServer="https://eip.kmc.gov.tw/api/",c["a"].prototype.$fileServer="https://eip.kmc.gov.tw/",c["a"].component("text-highlight",L["a"]);var N=t("bc3a"),$=t.n(N);c["a"].prototype.$axios=$.a;var q="";function B(){return M.apply(this,arguments)}function M(){return M=u()(regeneratorRuntime.mark((function e(){var n,t,r,o,a,u,i,l,f;return regeneratorRuntime.wrap((function(e){while(1)switch(e.prev=e.next){case 0:return e.next=2,C();case 2:n=e.sent,t=n.app,r=n.store,o=n.router,a=!1,u=function(e){a=!0;var n=Object(e)===e?o.resolve(e).route.fullPath:e;window.location.href=n},i=window.location.href.replace(window.location.origin,""),l=[void 0,void 0],f=0;case 11:if(!(!1===a&&f<l.length)){e.next=29;break}if("function"===typeof l[f]){e.next=14;break}return e.abrupt("continue",26);case 14:return e.prev=14,e.next=17,l[f]({app:t,router:o,store:r,Vue:c["a"],ssrContext:null,redirect:u,urlPath:i,publicPath:q});case 17:e.next=26;break;case 19:if(e.prev=19,e.t0=e["catch"](14),!e.t0||!e.t0.url){e.next=24;break}return window.location.href=e.t0.url,e.abrupt("return");case 24:return console.error("[Quasar] boot error:",e.t0),e.abrupt("return");case 26:f++,e.next=11;break;case 29:if(!0!==a){e.next=31;break}return e.abrupt("return");case 31:new c["a"](t);case 32:case"end":return e.stop()}}),e,null,[[14,19]])}))),M.apply(this,arguments)}B()},"31cd":function(e,n,t){},c091:function(e,n){}});