/* PowerPoint specific API library */
/* Version: 16.0.9009.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/


/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

var __extends=this&&this.__extends||function(b,a){for(var c in a)if(a.hasOwnProperty(c))b[c]=a[c];function d(){this.constructor=b}b.prototype=a===null?Object.create(a):(d.prototype=a.prototype,new d)},OfficeExt;(function(b){var a=function(){var a=true;function b(){}b.prototype.isMsAjaxLoaded=function(){var b="function",c="undefined";if(typeof Sys!==c&&typeof Type!==c&&Sys.StringBuilder&&typeof Sys.StringBuilder===b&&Type.registerNamespace&&typeof Type.registerNamespace===b&&Type.registerClass&&typeof Type.registerClass===b&&typeof Function._validateParams===b&&Sys.Serialization&&Sys.Serialization.JavaScriptSerializer&&typeof Sys.Serialization.JavaScriptSerializer.serialize===b)return a;else return false};b.prototype.loadMsAjaxFull=function(b){var a=(window.location.protocol.toLowerCase()==="https:"?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(a,b)};Object.defineProperty(b.prototype,"msAjaxError",{"get":function(){var a=this;if(a._msAjaxError==null&&a.isMsAjaxLoaded())a._msAjaxError=Error;return a._msAjaxError},"set":function(a){this._msAjaxError=a},enumerable:a,configurable:a});Object.defineProperty(b.prototype,"msAjaxString",{"get":function(){var a=this;if(a._msAjaxString==null&&a.isMsAjaxLoaded())a._msAjaxString=String;return a._msAjaxString},"set":function(a){this._msAjaxString=a},enumerable:a,configurable:a});Object.defineProperty(b.prototype,"msAjaxDebug",{"get":function(){var a=this;if(a._msAjaxDebug==null&&a.isMsAjaxLoaded())a._msAjaxDebug=Sys.Debug;return a._msAjaxDebug},"set":function(a){this._msAjaxDebug=a},enumerable:a,configurable:a});return b}();b.MicrosoftAjaxFactory=a})(OfficeExt||(OfficeExt={}));var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory,OSF=OSF||{},OfficeExt;(function(b){var a=function(){function a(a){this._internalStorage=a}a.prototype.getItem=function(a){try{return this._internalStorage&&this._internalStorage.getItem(a)}catch(b){return null}};a.prototype.setItem=function(b,a){try{this._internalStorage&&this._internalStorage.setItem(b,a)}catch(c){}};a.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(a){}};a.prototype.removeItem=function(a){try{this._internalStorage&&this._internalStorage.removeItem(a)}catch(b){}};a.prototype.getKeysWithPrefix=function(d){var b=[];try{for(var e=this._internalStorage&&this._internalStorage.length||0,a=0;a<e;a++){var c=this._internalStorage.key(a);c.indexOf(d)===0&&b.push(c)}}catch(f){}return b};return a}();b.SafeStorage=a})(OfficeExt||(OfficeExt={}));OSF.XdmFieldName={ConversationUrl:"ConversationUrl",AppId:"AppId"};OSF.WindowNameItemKeys={BaseFrameName:"baseFrameName",HostInfo:"hostInfo",XdmInfo:"xdmInfo",SerializerVersion:"serializerVersion",AppContext:"appContext"};OSF.OUtil=function(){var h="focus",g="on",n="configurable",m="writable",f="enumerable",e="",i="undefined",d=false,b=true,j=2147483647,a=null,c=-1,t=c,y="&_xdm_Info=",w="&_serializer_version=",x="_xdm_",B="_serializer_version=",p="#",v="&",k="class",s={},A=3e4,o=a,r=a,l=(new Date).getTime();function z(){var a=j*Math.random();a^=l^(new Date).getMilliseconds()<<Math.floor(Math.random()*(31-10));return a.toString(16)}function q(){if(!o){try{var b=window.sessionStorage}catch(c){b=a}o=new OfficeExt.SafeStorage(b)}return o}function u(e){for(var c=[],b=[],f=e.length,a,d=0;d<f;d++){a=e[d];if(a.tabIndex)if(a.tabIndex>0)b.push(a);else a.tabIndex===0&&c.push(a);else c.push(a)}b=b.sort(function(d,c){var a=d.tabIndex-c.tabIndex;if(a===0)a=b.indexOf(d)-b.indexOf(c);return a});return [].concat(b,c)}return {set_entropy:function(a){if(typeof a=="string")for(var b=0;b<a.length;b+=4){for(var d=0,c=0;c<4&&b+c<a.length;c++)d=(d<<8)+a.charCodeAt(b+c);l^=d}else if(typeof a=="number")l^=a;else l^=j*Math.random();l&=j},extend:function(b,a){var c=function(){};c.prototype=a.prototype;b.prototype=new c;b.prototype.constructor=b;b.uber=a.prototype;if(a.prototype.constructor===Object.prototype.constructor)a.prototype.constructor=a},setNamespace:function(b,a){if(a&&b&&!a[b])a[b]={}},unsetNamespace:function(b,a){if(a&&b&&a[b])delete a[b]},serializeSettings:function(b){var d={};for(var c in b){var a=b[c];try{if(JSON)a=JSON.stringify(a,function(a,b){return OSF.OUtil.isDate(this[a])?OSF.DDA.SettingsManager.DateJSONPrefix+this[a].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix:b});else a=Sys.Serialization.JavaScriptSerializer.serialize(a);d[c]=a}catch(e){}}return d},deserializeSettings:function(d){var f={};d=d||{};for(var e in d){var a=d[e];try{if(JSON)a=JSON.parse(a,function(d,a){var b;if(typeof a==="string"&&a&&a.length>6&&a.slice(0,5)===OSF.DDA.SettingsManager.DateJSONPrefix&&a.slice(c)===OSF.DDA.SettingsManager.DataJSONSuffix){b=new Date(parseInt(a.slice(5,c)));if(b)return b}return a});else a=Sys.Serialization.JavaScriptSerializer.deserialize(a,b);f[e]=a}catch(g){}}return f},loadScript:function(f,g,h){if(f&&g){var k=window.document,c=s[f];if(!c){var e=k.createElement("script");e.type="text/javascript";c={loaded:d,pendingCallbacks:[g],timer:a};s[f]=c;var i=function(){if(c.timer!=a){clearTimeout(c.timer);delete c.timer}c.loaded=b;for(var e=c.pendingCallbacks.length,d=0;d<e;d++){var f=c.pendingCallbacks.shift();f()}},j=function(){delete s[f];if(c.timer!=a){clearTimeout(c.timer);delete c.timer}for(var d=c.pendingCallbacks.length,b=0;b<d;b++){var e=c.pendingCallbacks.shift();e()}};if(e.readyState)e.onreadystatechange=function(){if(e.readyState=="loaded"||e.readyState=="complete"){e.onreadystatechange=a;i()}};else e.onload=i;e.onerror=j;h=h||A;c.timer=setTimeout(j,h);e.setAttribute("crossOrigin","anonymous");e.src=f;k.getElementsByTagName("head")[0].appendChild(e)}else if(c.loaded)g();else c.pendingCallbacks.push(g)}},loadCSS:function(c){if(c){var b=window.document,a=b.createElement("link");a.type="text/css";a.rel="stylesheet";a.href=c;b.getElementsByTagName("head")[0].appendChild(a)}},parseEnum:function(b,c){var a=c[b.trim()];if(typeof a==i){OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+b);throw OsfMsAjaxFactory.msAjaxError.argument("str")}return a},delayExecutionAndCache:function(){var a={calc:arguments[0]};return function(){if(a.calc){a.val=a.calc.apply(this,arguments);delete a.calc}return a.val}},getUniqueId:function(){t=t+1;return t.toString()},formatString:function(){var a=arguments,b=a[0];return b.replace(/{(\d+)}/gm,function(d,b){var c=parseInt(b,10)+1;return a[c]===undefined?"{"+b+"}":a[c]})},generateConversationId:function(){return [z(),z(),(new Date).getTime().toString()].join("_")},getFrameName:function(a){return x+a+this.generateConversationId()},addXdmInfoAsHash:function(b,a){return OSF.OUtil.addInfoAsHash(b,y,a,d)},addSerializerVersionAsHash:function(c,a){return OSF.OUtil.addInfoAsHash(c,w,a,b)},addInfoAsHash:function(b,g,c,i){b=b.trim()||e;var f=b.split(p),h=f.shift(),d=f.join(p),a;if(i)a=[g,encodeURIComponent(c),d].join(e);else a=[d,g,c].join(e);return [h,p,a].join(e)},parseHostInfoFromWindowName:function(a,b){return OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.HostInfo)},parseXdmInfo:function(b){var a=OSF.OUtil.parseXdmInfoWithGivenFragment(b,window.location.hash);if(!a)a=OSF.OUtil.parseXdmInfoFromWindowName(b,window.name);return a},parseXdmInfoFromWindowName:function(a,b){return OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.XdmInfo)},parseXdmInfoWithGivenFragment:function(a,b){return OSF.OUtil.parseInfoWithGivenFragment(y,x,d,a,b)},parseSerializerVersion:function(b){var a=OSF.OUtil.parseSerializerVersionWithGivenFragment(b,window.location.hash);if(isNaN(a))a=OSF.OUtil.parseSerializerVersionFromWindowName(b,window.name);return a},parseSerializerVersionFromWindowName:function(a,b){return parseInt(OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.SerializerVersion))},parseSerializerVersionWithGivenFragment:function(a,c){return parseInt(OSF.OUtil.parseInfoWithGivenFragment(w,B,b,a,c))},parseInfoFromWindowName:function(g,h,f){try{var b=JSON.parse(h),c=b!=a?b[f]:a,d=q();if(!g&&d&&b!=a){var e=b[OSF.WindowNameItemKeys.BaseFrameName]+f;if(c)d.setItem(e,c);else c=d.getItem(e)}return c}catch(i){return a}},parseInfoWithGivenFragment:function(m,j,k,i,l){var f=l.split(m),b=f.length>1?f[f.length-1]:a;if(k&&b!=a){if(b.indexOf(v)>=0)b=b.split(v)[0];b=decodeURIComponent(b)}var d=q();if(!i&&d){var e=window.name.indexOf(j);if(e>c){var g=window.name.indexOf(";",e);if(g==c)g=window.name.length;var h=window.name.substring(e,g);if(b)d.setItem(h,b);else b=d.getItem(h)}}return b},getConversationId:function(){var c=window.location.search,b=a;if(c){var d=c.indexOf("&");b=d>0?c.substring(1,d):c.substr(1);if(b&&b.charAt(b.length-1)==="="){b=b.substring(0,b.length-1);if(b)b=decodeURIComponent(b)}}return b},getInfoItems:function(b){var a=b.split("$");if(typeof a[1]==i)a=b.split("|");if(typeof a[1]==i)a=b.split("%7C");return a},getXdmFieldValue:function(f,d){var b=e,c=OSF.OUtil.parseXdmInfo(d);if(c){var a=OSF.OUtil.getInfoItems(c);if(a!=undefined&&a.length>=3)switch(f){case OSF.XdmFieldName.ConversationUrl:b=a[2];break;case OSF.XdmFieldName.AppId:b=a[1]}}return b},validateParamObject:function(f,e){var a=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:d},{name:"expectedProperties",type:Object,mayBeNull:d},{name:"callback",type:Function,mayBeNull:b}]);if(a)throw a;for(var c in e){a=Function._validateParameter(f[c],e[c],c);if(a)throw a}},writeProfilerMark:function(a){if(window.msWriteProfilerMark){window.msWriteProfilerMark(a);OsfMsAjaxFactory.msAjaxDebug.trace(a)}},outputDebug:function(a){typeof OsfMsAjaxFactory!==i&&OsfMsAjaxFactory.msAjaxDebug&&OsfMsAjaxFactory.msAjaxDebug.trace&&OsfMsAjaxFactory.msAjaxDebug.trace(a)},defineNondefaultProperty:function(e,f,a,c){a=a||{};for(var g in c){var d=c[g];if(a[d]==undefined)a[d]=b}Object.defineProperty(e,f,a);return e},defineNondefaultProperties:function(c,a,d){a=a||{};for(var b in a)OSF.OUtil.defineNondefaultProperty(c,b,a[b],d);return c},defineEnumerableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,[f])},defineEnumerableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,[f])},defineMutableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,[m,f,n])},defineMutableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,[m,f,n])},finalizeProperties:function(e,c){c=c||{};for(var g=Object.getOwnPropertyNames(e),i=g.length,f=0;f<i;f++){var h=g[f],a=Object.getOwnPropertyDescriptor(e,h);if(!a.get&&!a.set)a.writable=c.writable||d;a.configurable=c.configurable||d;a.enumerable=c.enumerable||b;Object.defineProperty(e,h,a)}return e},mapList:function(a,c){var b=[];if(a)for(var d in a)b.push(c(a[d]));return b},listContainsKey:function(c,e){for(var a in c)if(e==a)return b;return d},listContainsValue:function(a,c){for(var e in a)if(c==a[e])return b;return d},augmentList:function(a,b){var d=a.push?function(c,b){a.push(b)}:function(c,b){a[c]=b};for(var c in b)d(c,b[c])},redefineList:function(a,b){for(var d in a)delete a[d];for(var c in b)a[c]=b[c]},isArray:function(a){return Object.prototype.toString.apply(a)==="[object Array]"},isFunction:function(a){return Object.prototype.toString.apply(a)==="[object Function]"},isDate:function(a){return Object.prototype.toString.apply(a)==="[object Date]"},addEventListener:function(a,b,c){if(a.addEventListener)a.addEventListener(b,c,d);else if(Sys.Browser.agent===Sys.Browser.InternetExplorer&&a.attachEvent)a.attachEvent(g+b,c);else a[g+b]=c},removeEventListener:function(b,c,e){if(b.removeEventListener)b.removeEventListener(c,e,d);else if(Sys.Browser.agent===Sys.Browser.InternetExplorer&&b.detachEvent)b.detachEvent(g+c,e);else b[g+c]=a},getCookieValue:function(b){var a=RegExp(b+"[^;]+").exec(document.cookie);return a.toString().replace(/^[^=]+./,e)},xhrGet:function(f,e,c){var a;try{a=new XMLHttpRequest;a.onreadystatechange=function(){if(a.readyState==4)if(a.status==200)e(a.responseText);else c(a.status)};a.open("GET",f,b);a.send()}catch(d){c(d)}},xhrGetFull:function(h,f,g,c){var a,e=f;try{a=new XMLHttpRequest;a.onreadystatechange=function(){if(a.readyState==4)if(a.status==200)g(a,e);else c(a.status)};a.open("GET",h,b);a.send()}catch(d){c(d)}},encodeBase64:function(c){if(!c)return c;var o="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",m=[],b=[],i=0,k,h,j,d,f,g,a,n=c.length;do{k=c.charCodeAt(i++);h=c.charCodeAt(i++);j=c.charCodeAt(i++);a=0;d=k&255;f=k>>8;g=h&255;b[a++]=d>>2;b[a++]=(d&3)<<4|f>>4;b[a++]=(f&15)<<2|g>>6;b[a++]=g&63;if(!isNaN(h)){d=h>>8;f=j&255;g=j>>8;b[a++]=d>>2;b[a++]=(d&3)<<4|f>>4;b[a++]=(f&15)<<2|g>>6;b[a++]=g&63}if(isNaN(h))b[a-1]=64;else if(isNaN(j)){b[a-2]=64;b[a-1]=64}for(var l=0;l<a;l++)m.push(o.charAt(b[l]))}while(i<n);return m.join(e)},getSessionStorage:function(){return q()},getLocalStorage:function(){if(!r){try{var b=window.localStorage}catch(c){b=a}r=new OfficeExt.SafeStorage(b)}return r},convertIntToCssHexColor:function(b){var a="#"+(Number(b)+16777216).toString(16).slice(-6);return a},attachClickHandler:function(a,b){a.onclick=function(){b()};a.ontouchend=function(a){b();a.preventDefault()}},getQueryStringParamValue:function(a,c){var f=Function._validateParams(arguments,[{name:"queryString",type:String,mayBeNull:d},{name:"paramName",type:String,mayBeNull:d}]);if(f){OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");return e}var b=new RegExp("[\\?&]"+c+"=([^&#]*)","i");if(!b.test(a)){OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");return e}return b.exec(a)[1]},isiOS:function(){return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g)?b:d},isChrome:function(){return window.navigator.userAgent.indexOf("Chrome")>0&&!OSF.OUtil.isEdge()},isEdge:function(){return window.navigator.userAgent.indexOf("Edge")>0},isIE:function(){return window.navigator.userAgent.indexOf("Trident")>0},isFirefox:function(){return window.navigator.userAgent.indexOf("Firefox")>0},shallowCopy:function(b){if(b==a)return a;else if(!(b instanceof Object))return b;else if(Array.isArray(b)){for(var e=[],d=0;d<b.length;d++)e.push(b[d]);return e}else{var f=b.constructor();for(var c in b)if(b.hasOwnProperty(c))f[c]=b[c];return f}},createObject:function(b){var d=a;if(b){d={};for(var e=b.length,c=0;c<e;c++)d[b[c].name]=b[c].value}return d},addClass:function(a,b){if(!OSF.OUtil.hasClass(a,b)){var c=a.getAttribute(k);if(c)a.setAttribute(k,c+" "+b);else a.setAttribute(k,b)}},removeClass:function(b,c){if(OSF.OUtil.hasClass(b,c)){var a=b.getAttribute(k),d=new RegExp("(\\s|^)"+c+"(\\s|$)");a=a.replace(d,e);b.setAttribute(k,a)}},hasClass:function(c,b){var a=c.getAttribute(k);return a&&a.match(new RegExp("(\\s|^)"+b+"(\\s|$)"))},focusToFirstTabbable:function(e,j){var g,i=d,f,k=function(){i=b},l=function(d,a,b){if(a<0||a>d)return c;else if(a===0&&b)return c;else if(a===d-1&&!b)return c;if(b)return a-1;else return a+1};e=u(e);g=j?e.length-1:0;if(e.length===0)return a;while(!i&&g>=0&&g<e.length){f=e[g];window.focus();f.addEventListener(h,k);f.focus();f.removeEventListener(h,k);g=l(e.length,g,j);if(!i&&f===document.activeElement)i=b}if(i)return f;else return a},focusToNextTabbable:function(f,o,m){var k,e,i=d,g,l=function(){i=b},n=function(b,d){for(var a=0;a<b.length;a++)if(b[a]===d)return a;return c},j=function(d,a,b){if(a<0||a>d)return c;else if(a===0&&b)return c;else if(a===d-1&&!b)return c;if(b)return a-1;else return a+1};f=u(f);k=n(f,o);e=j(f.length,k,m);if(e<0)return a;while(!i&&e>=0&&e<f.length){g=f[e];g.addEventListener(h,l);g.focus();g.removeEventListener(h,l);e=j(f.length,e,m);if(!i&&g===document.activeElement)i=b}if(i)return g;else return a}}}();OSF.OUtil.Guid=function(){var a=["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f"];return {generateNewGuid:function(){for(var c="",d=(new Date).getTime(),b=0;b<32&&d>0;b++){if(b==8||b==12||b==16||b==20)c+="-";c+=a[d%16];d=Math.floor(d/16)}for(;b<32;b++){if(b==8||b==12||b==16||b==20)c+="-";c+=a[Math.floor(Math.random()*16)]}return c}}}();window.OSF=OSF;OSF.OUtil.setNamespace("OSF",window);OSF.MessageIDs={FetchBundleUrl:0,LoadReactBundle:1,LoadBundleSuccess:2,LoadBundleError:3};OSF.AppName={Unsupported:0,Excel:1,Word:2,PowerPoint:4,Outlook:8,ExcelWebApp:16,WordWebApp:32,OutlookWebApp:64,Project:128,AccessWebApp:256,PowerpointWebApp:512,ExcelIOS:1024,Sway:2048,WordIOS:4096,PowerPointIOS:8192,Access:16384,Lync:32768,OutlookIOS:65536,OneNoteWebApp:131072,OneNote:262144,ExcelWinRT:524288,WordWinRT:1048576,PowerpointWinRT:2097152,OutlookAndroid:4194304,OneNoteWinRT:8388608,ExcelAndroid:8388609,VisioWebApp:8388610,OneNoteIOS:8388611,WordAndroid:8388613,PowerpointAndroid:8388614};OSF.InternalPerfMarker={DataCoercionBegin:"Agave.HostCall.CoerceDataStart",DataCoercionEnd:"Agave.HostCall.CoerceDataEnd"};OSF.HostCallPerfMarker={IssueCall:"Agave.HostCall.IssueCall",ReceiveResponse:"Agave.HostCall.ReceiveResponse",RuntimeExceptionRaised:"Agave.HostCall.RuntimeExecptionRaised"};OSF.AgaveHostAction={Select:0,UnSelect:1,CancelDialog:2,InsertAgave:3,CtrlF6In:4,CtrlF6Exit:5,CtrlF6ExitShift:6,SelectWithError:7,NotifyHostError:8,RefreshAddinCommands:9,PageIsReady:10,TabIn:11,TabInShift:12,TabExit:13,TabExitShift:14,EscExit:15,F2Exit:16,ExitNoFocusable:17,ExitNoFocusableShift:18,MouseEnter:19,MouseLeave:20};OSF.SharedConstants={NotificationConversationIdSuffix:"_ntf"};OSF.DialogMessageType={DialogMessageReceived:0,DialogParentMessageReceived:1,DialogClosed:12006};OSF.OfficeAppContext=function(y,u,p,n,r,v,q,t,x,j,w,l,k,m,h,g,f,e,i,c,d,s,o,b){var a=this;a._id=y;a._appName=u;a._appVersion=p;a._appUILocale=n;a._dataLocale=r;a._docUrl=v;a._clientMode=q;a._settings=t;a._reason=x;a._osfControlType=j;a._eToken=w;a._correlationId=l;a._appInstanceId=k;a._touchEnabled=m;a._commerceAllowed=h;a._appMinorVersion=g;a._requirementMatrix=f;a._hostCustomMessage=e;a._hostFullVersion=i;a._isDialog=false;a._clientWindowHeight=c;a._clientWindowWidth=d;a._addinName=s;a._appDomains=o;a._dialogRequirementMatrix=b;a.get_id=function(){return this._id};a.get_appName=function(){return this._appName};a.get_appVersion=function(){return this._appVersion};a.get_appUILocale=function(){return this._appUILocale};a.get_dataLocale=function(){return this._dataLocale};a.get_docUrl=function(){return this._docUrl};a.get_clientMode=function(){return this._clientMode};a.get_bindings=function(){return this._bindings};a.get_settings=function(){return this._settings};a.get_reason=function(){return this._reason};a.get_osfControlType=function(){return this._osfControlType};a.get_eToken=function(){return this._eToken};a.get_correlationId=function(){return this._correlationId};a.get_appInstanceId=function(){return this._appInstanceId};a.get_touchEnabled=function(){return this._touchEnabled};a.get_commerceAllowed=function(){return this._commerceAllowed};a.get_appMinorVersion=function(){return this._appMinorVersion};a.get_requirementMatrix=function(){return this._requirementMatrix};a.get_dialogRequirementMatrix=function(){return this._dialogRequirementMatrix};a.get_hostCustomMessage=function(){return this._hostCustomMessage};a.get_hostFullVersion=function(){return this._hostFullVersion};a.get_isDialog=function(){return this._isDialog};a.get_clientWindowHeight=function(){return this._clientWindowHeight};a.get_clientWindowWidth=function(){return this._clientWindowWidth};a.get_addinName=function(){return this._addinName};a.get_appDomains=function(){return this._appDomains}};OSF.OsfControlType={DocumentLevel:0,ContainerLevel:1};OSF.ClientMode={ReadOnly:0,ReadWrite:1};OSF.OUtil.setNamespace("Microsoft",window);OSF.OUtil.setNamespace("Office",Microsoft);OSF.OUtil.setNamespace("Client",Microsoft.Office);OSF.OUtil.setNamespace("WebExtension",Microsoft.Office);Microsoft.Office.WebExtension.InitializationReason={Inserted:"inserted",DocumentOpened:"documentOpened"};Microsoft.Office.WebExtension.ValueFormat={Unformatted:"unformatted",Formatted:"formatted"};Microsoft.Office.WebExtension.FilterType={All:"all"};Microsoft.Office.WebExtension.PlatformType={PC:"PC",OfficeOnline:"OfficeOnline",Mac:"Mac",iOS:"iOS",Android:"Android",Universal:"Universal"};Microsoft.Office.WebExtension.HostType={Word:"Word",Excel:"Excel",PowerPoint:"PowerPoint",Outlook:"Outlook",OneNote:"OneNote",Project:"Project",Access:"Access"};Microsoft.Office.WebExtension.Parameters={BindingType:"bindingType",CoercionType:"coercionType",ValueFormat:"valueFormat",FilterType:"filterType",Columns:"columns",SampleData:"sampleData",GoToType:"goToType",SelectionMode:"selectionMode",Id:"id",PromptText:"promptText",ItemName:"itemName",FailOnCollision:"failOnCollision",StartRow:"startRow",StartColumn:"startColumn",RowCount:"rowCount",ColumnCount:"columnCount",Callback:"callback",AsyncContext:"asyncContext",Data:"data",Rows:"rows",OverwriteIfStale:"overwriteIfStale",FileType:"fileType",EventType:"eventType",Handler:"handler",SliceSize:"sliceSize",SliceIndex:"sliceIndex",ActiveView:"activeView",Status:"status",PlatformType:"platformType",HostType:"hostType",ForceConsent:"forceConsent",ForceAddAccount:"forceAddAccount",AuthChallenge:"authChallenge",Reserved:"reserved",Xml:"xml",Namespace:"namespace",Prefix:"prefix",XPath:"xPath",Text:"text",ImageLeft:"imageLeft",ImageTop:"imageTop",ImageWidth:"imageWidth",ImageHeight:"imageHeight",TaskId:"taskId",FieldId:"fieldId",FieldValue:"fieldValue",ServerUrl:"serverUrl",ListName:"listName",ResourceId:"resourceId",ViewType:"viewType",ViewName:"viewName",GetRawValue:"getRawValue",CellFormat:"cellFormat",TableOptions:"tableOptions",TaskIndex:"taskIndex",ResourceIndex:"resourceIndex",CustomFieldId:"customFieldId",Url:"url",MessageHandler:"messageHandler",Width:"width",Height:"height",RequireHTTPs:"requireHTTPS",MessageToParent:"messageToParent",DisplayInIframe:"displayInIframe",MessageContent:"messageContent",HideTitle:"hideTitle",UseDeviceIndependentPixels:"useDeviceIndependentPixels",AppCommandInvocationCompletedData:"appCommandInvocationCompletedData",Base64:"base64",FormId:"formId"};OSF.OUtil.setNamespace("DDA",OSF);OSF.DDA.DocumentMode={ReadOnly:1,ReadWrite:0};OSF.DDA.PropertyDescriptors={AsyncResultStatus:"AsyncResultStatus"};OSF.DDA.EventDescriptors={};OSF.DDA.ListDescriptors={};OSF.DDA.UI={};OSF.DDA.getXdmEventName=function(b,a){if(a==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged||a==Microsoft.Office.WebExtension.EventType.BindingDataChanged||a==Microsoft.Office.WebExtension.EventType.DataNodeDeleted||a==Microsoft.Office.WebExtension.EventType.DataNodeInserted||a==Microsoft.Office.WebExtension.EventType.DataNodeReplaced)return b+"_"+a;else return a};OSF.DDA.MethodDispId={dispidMethodMin:64,dispidGetSelectedDataMethod:64,dispidSetSelectedDataMethod:65,dispidAddBindingFromSelectionMethod:66,dispidAddBindingFromPromptMethod:67,dispidGetBindingMethod:68,dispidReleaseBindingMethod:69,dispidGetBindingDataMethod:70,dispidSetBindingDataMethod:71,dispidAddRowsMethod:72,dispidClearAllRowsMethod:73,dispidGetAllBindingsMethod:74,dispidLoadSettingsMethod:75,dispidSaveSettingsMethod:76,dispidGetDocumentCopyMethod:77,dispidAddBindingFromNamedItemMethod:78,dispidAddColumnsMethod:79,dispidGetDocumentCopyChunkMethod:80,dispidReleaseDocumentCopyMethod:81,dispidNavigateToMethod:82,dispidGetActiveViewMethod:83,dispidGetDocumentThemeMethod:84,dispidGetOfficeThemeMethod:85,dispidGetFilePropertiesMethod:86,dispidClearFormatsMethod:87,dispidSetTableOptionsMethod:88,dispidSetFormatsMethod:89,dispidExecuteRichApiRequestMethod:93,dispidAppCommandInvocationCompletedMethod:94,dispidCloseContainerMethod:97,dispidGetAccessTokenMethod:98,dispidOpenBrowserWindow:102,dispidCreateDocumentMethod:105,dispidInsertFormMethod:106,dispidGetSelectedTaskMethod:110,dispidGetSelectedResourceMethod:111,dispidGetTaskMethod:112,dispidGetResourceFieldMethod:113,dispidGetWSSUrlMethod:114,dispidGetTaskFieldMethod:115,dispidGetProjectFieldMethod:116,dispidGetSelectedViewMethod:117,dispidGetTaskByIndexMethod:118,dispidGetResourceByIndexMethod:119,dispidSetTaskFieldMethod:120,dispidSetResourceFieldMethod:121,dispidGetMaxTaskIndexMethod:122,dispidGetMaxResourceIndexMethod:123,dispidCreateTaskMethod:124,dispidAddDataPartMethod:128,dispidGetDataPartByIdMethod:129,dispidGetDataPartsByNamespaceMethod:130,dispidGetDataPartXmlMethod:131,dispidGetDataPartNodesMethod:132,dispidDeleteDataPartMethod:133,dispidGetDataNodeValueMethod:134,dispidGetDataNodeXmlMethod:135,dispidGetDataNodesMethod:136,dispidSetDataNodeValueMethod:137,dispidSetDataNodeXmlMethod:138,dispidAddDataNamespaceMethod:139,dispidGetDataUriByPrefixMethod:140,dispidGetDataPrefixByUriMethod:141,dispidGetDataNodeTextMethod:142,dispidSetDataNodeTextMethod:143,dispidMessageParentMethod:144,dispidSendMessageMethod:145,dispidMethodMax:145};OSF.DDA.EventDispId={dispidEventMin:0,dispidInitializeEvent:0,dispidSettingsChangedEvent:1,dispidDocumentSelectionChangedEvent:2,dispidBindingSelectionChangedEvent:3,dispidBindingDataChangedEvent:4,dispidDocumentOpenEvent:5,dispidDocumentCloseEvent:6,dispidActiveViewChangedEvent:7,dispidDocumentThemeChangedEvent:8,dispidOfficeThemeChangedEvent:9,dispidDialogMessageReceivedEvent:10,dispidDialogNotificationShownInAddinEvent:11,dispidDialogParentMessageReceivedEvent:12,dispidObjectDeletedEvent:13,dispidObjectSelectionChangedEvent:14,dispidObjectDataChangedEvent:15,dispidContentControlAddedEvent:16,dispidActivationStatusChangedEvent:32,dispidRichApiMessageEvent:33,dispidAppCommandInvokedEvent:39,dispidOlkItemSelectedChangedEvent:46,dispidOlkRecipientsChangedEvent:47,dispidOlkAppointmentTimeChangedEvent:48,dispidTaskSelectionChangedEvent:56,dispidResourceSelectionChangedEvent:57,dispidViewSelectionChangedEvent:58,dispidDataNodeAddedEvent:60,dispidDataNodeReplacedEvent:61,dispidDataNodeDeletedEvent:62,dispidEventMax:63};OSF.DDA.ErrorCodeManager=function(){var a={};return {getErrorArgs:function(c){var b=a[c];if(!b)b=a[this.errorCodes.ooeInternalError];else{if(!b.name)b.name=a[this.errorCodes.ooeInternalError].name;if(!b.message)b.message=a[this.errorCodes.ooeInternalError].message}return b},addErrorMessage:function(c,b){a[c]=b},errorCodes:{ooeSuccess:0,ooeChunkResult:1,ooeCoercionTypeNotSupported:1e3,ooeGetSelectionNotMatchDataType:1001,ooeCoercionTypeNotMatchBinding:1002,ooeInvalidGetRowColumnCounts:1003,ooeSelectionNotSupportCoercionType:1004,ooeInvalidGetStartRowColumn:1005,ooeNonUniformPartialGetNotSupported:1006,ooeGetDataIsTooLarge:1008,ooeFileTypeNotSupported:1009,ooeGetDataParametersConflict:1010,ooeInvalidGetColumns:1011,ooeInvalidGetRows:1012,ooeInvalidReadForBlankRow:1013,ooeUnsupportedDataObject:2e3,ooeCannotWriteToSelection:2001,ooeDataNotMatchSelection:2002,ooeOverwriteWorksheetData:2003,ooeDataNotMatchBindingSize:2004,ooeInvalidSetStartRowColumn:2005,ooeInvalidDataFormat:2006,ooeDataNotMatchCoercionType:2007,ooeDataNotMatchBindingType:2008,ooeSetDataIsTooLarge:2009,ooeNonUniformPartialSetNotSupported:2010,ooeInvalidSetColumns:2011,ooeInvalidSetRows:2012,ooeSetDataParametersConflict:2013,ooeCellDataAmountBeyondLimits:2014,ooeSelectionCannotBound:3e3,ooeBindingNotExist:3002,ooeBindingToMultipleSelection:3003,ooeInvalidSelectionForBindingType:3004,ooeOperationNotSupportedOnThisBindingType:3005,ooeNamedItemNotFound:3006,ooeMultipleNamedItemFound:3007,ooeInvalidNamedItemForBindingType:3008,ooeUnknownBindingType:3009,ooeOperationNotSupportedOnMatrixData:3010,ooeInvalidColumnsForBinding:3011,ooeSettingNameNotExist:4e3,ooeSettingsCannotSave:4001,ooeSettingsAreStale:4002,ooeOperationNotSupported:5e3,ooeInternalError:5001,ooeDocumentReadOnly:5002,ooeEventHandlerNotExist:5003,ooeInvalidApiCallInContext:5004,ooeShuttingDown:5005,ooeUnsupportedEnumeration:5007,ooeIndexOutOfRange:5008,ooeBrowserAPINotSupported:5009,ooeInvalidParam:5010,ooeRequestTimeout:5011,ooeInvalidOrTimedOutSession:5012,ooeInvalidApiArguments:5013,ooeTooManyIncompleteRequests:5100,ooeRequestTokenUnavailable:5101,ooeActivityLimitReached:5102,ooeCustomXmlNodeNotFound:6e3,ooeCustomXmlError:6100,ooeCustomXmlExceedQuota:6101,ooeCustomXmlOutOfDate:6102,ooeNoCapability:7e3,ooeCannotNavTo:7001,ooeSpecifiedIdNotExist:7002,ooeNavOutOfBound:7004,ooeElementMissing:8e3,ooeProtectedError:8001,ooeInvalidCellsValue:8010,ooeInvalidTableOptionValue:8011,ooeInvalidFormatValue:8012,ooeRowIndexOutOfRange:8020,ooeColIndexOutOfRange:8021,ooeFormatValueOutOfRange:8022,ooeCellFormatAmountBeyondLimits:8023,ooeMemoryFileLimit:11000,ooeNetworkProblemRetrieveFile:11001,ooeInvalidSliceSize:11002,ooeInvalidCallback:11101,ooeInvalidWidth:12000,ooeInvalidHeight:12001,ooeNavigationError:12002,ooeInvalidScheme:12003,ooeAppDomains:12004,ooeRequireHTTPS:12005,ooeWebDialogClosed:12006,ooeDialogAlreadyOpened:12007,ooeEndUserAllow:12008,ooeEndUserIgnore:12009,ooeNotUILessDialog:12010,ooeCrossZone:12011,ooeNotSSOAgave:13000,ooeSSOUserNotSignedIn:13001,ooeSSOUserAborted:13002,ooeSSOUnsupportedUserIdentity:13003,ooeSSOInvalidResourceUrl:13004,ooeSSOInvalidGrant:13005,ooeSSOClientError:13006,ooeSSOServerError:13007,ooeAddinIsAlreadyRequestingToken:13008,ooeSSOUserConsentNotSupportedByCurrentAddinCategory:13009,ooeSSOConnectionLost:13010},initializeErrorMessages:function(b){a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported]={name:b.L_InvalidCoercion,message:b.L_CoercionTypeNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType]={name:b.L_DataReadError,message:b.L_GetSelectionNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding]={name:b.L_InvalidCoercion,message:b.L_CoercionTypeNotMatchBinding};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts]={name:b.L_DataReadError,message:b.L_InvalidGetRowColumnCounts};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType]={name:b.L_DataReadError,message:b.L_SelectionNotSupportCoercionType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn]={name:b.L_DataReadError,message:b.L_InvalidGetStartRowColumn};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported]={name:b.L_DataReadError,message:b.L_NonUniformPartialGetNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge]={name:b.L_DataReadError,message:b.L_GetDataIsTooLarge};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported]={name:b.L_DataReadError,message:b.L_FileTypeNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict]={name:b.L_DataReadError,message:b.L_GetDataParametersConflict};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns]={name:b.L_DataReadError,message:b.L_InvalidGetColumns};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows]={name:b.L_DataReadError,message:b.L_InvalidGetRows};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow]={name:b.L_DataReadError,message:b.L_InvalidReadForBlankRow};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject]={name:b.L_DataWriteError,message:b.L_UnsupportedDataObject};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection]={name:b.L_DataWriteError,message:b.L_CannotWriteToSelection};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection]={name:b.L_DataWriteError,message:b.L_DataNotMatchSelection};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData]={name:b.L_DataWriteError,message:b.L_OverwriteWorksheetData};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize]={name:b.L_DataWriteError,message:b.L_DataNotMatchBindingSize};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn]={name:b.L_DataWriteError,message:b.L_InvalidSetStartRowColumn};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat]={name:b.L_InvalidFormat,message:b.L_InvalidDataFormat};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType]={name:b.L_InvalidDataObject,message:b.L_DataNotMatchCoercionType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType]={name:b.L_InvalidDataObject,message:b.L_DataNotMatchBindingType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge]={name:b.L_DataWriteError,message:b.L_SetDataIsTooLarge};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported]={name:b.L_DataWriteError,message:b.L_NonUniformPartialSetNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns]={name:b.L_DataWriteError,message:b.L_InvalidSetColumns};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows]={name:b.L_DataWriteError,message:b.L_InvalidSetRows};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict]={name:b.L_DataWriteError,message:b.L_SetDataParametersConflict};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound]={name:b.L_BindingCreationError,message:b.L_SelectionCannotBound};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist]={name:b.L_InvalidBindingError,message:b.L_BindingNotExist};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection]={name:b.L_BindingCreationError,message:b.L_BindingToMultipleSelection};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType]={name:b.L_BindingCreationError,message:b.L_InvalidSelectionForBindingType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType]={name:b.L_InvalidBindingOperation,message:b.L_OperationNotSupportedOnThisBindingType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound]={name:b.L_BindingCreationError,message:b.L_NamedItemNotFound};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound]={name:b.L_BindingCreationError,message:b.L_MultipleNamedItemFound};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType]={name:b.L_BindingCreationError,message:b.L_InvalidNamedItemForBindingType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType]={name:b.L_InvalidBinding,message:b.L_UnknownBindingType};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData]={name:b.L_InvalidBindingOperation,message:b.L_OperationNotSupportedOnMatrixData};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding]={name:b.L_InvalidBinding,message:b.L_InvalidColumnsForBinding};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist]={name:b.L_ReadSettingsError,message:b.L_SettingNameNotExist};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave]={name:b.L_SaveSettingsError,message:b.L_SettingsCannotSave};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale]={name:b.L_SettingsStaleError,message:b.L_SettingsAreStale};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported]={name:b.L_HostError,message:b.L_OperationNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError]={name:b.L_InternalError,message:b.L_InternalErrorDescription};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly]={name:b.L_PermissionDenied,message:b.L_DocumentReadOnly};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist]={name:b.L_EventRegistrationError,message:b.L_EventHandlerNotExist};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext]={name:b.L_InvalidAPICall,message:b.L_InvalidApiCallInContext};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown]={name:b.L_ShuttingDown,message:b.L_ShuttingDown};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration]={name:b.L_UnsupportedEnumeration,message:b.L_UnsupportedEnumerationMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange]={name:b.L_IndexOutOfRange,message:b.L_IndexOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported]={name:b.L_APINotSupported,message:b.L_BrowserAPINotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout]={name:b.L_APICallFailed,message:b.L_RequestTimeout};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession]={name:b.L_InvalidOrTimedOutSession,message:b.L_InvalidOrTimedOutSessionMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests]={name:b.L_APICallFailed,message:b.L_TooManyIncompleteRequests};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable]={name:b.L_APICallFailed,message:b.L_RequestTokenUnavailable};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached]={name:b.L_APICallFailed,message:b.L_ActivityLimitReached};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments]={name:b.L_APICallFailed,message:b.L_InvalidApiArgumentsMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound]={name:b.L_InvalidNode,message:b.L_CustomXmlNodeNotFound};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError]={name:b.L_CustomXmlError,message:b.L_CustomXmlError};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota]={name:b.L_CustomXmlExceedQuotaName,message:b.L_CustomXmlExceedQuotaMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate]={name:b.L_CustomXmlOutOfDateName,message:b.L_CustomXmlOutOfDateMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability]={name:b.L_PermissionDenied,message:b.L_NoCapability};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo]={name:b.L_CannotNavigateTo,message:b.L_CannotNavigateTo};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist]={name:b.L_SpecifiedIdNotExist,message:b.L_SpecifiedIdNotExist};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound]={name:b.L_NavOutOfBound,message:b.L_NavOutOfBound};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits]={name:b.L_DataWriteReminder,message:b.L_CellDataAmountBeyondLimits};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing]={name:b.L_MissingParameter,message:b.L_ElementMissing};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError]={name:b.L_PermissionDenied,message:b.L_NoCapability};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue]={name:b.L_InvalidValue,message:b.L_InvalidCellsValue};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue]={name:b.L_InvalidValue,message:b.L_InvalidTableOptionValue};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue]={name:b.L_InvalidValue,message:b.L_InvalidFormatValue};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange]={name:b.L_OutOfRange,message:b.L_RowIndexOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange]={name:b.L_OutOfRange,message:b.L_ColIndexOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange]={name:b.L_OutOfRange,message:b.L_FormatValueOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits]={name:b.L_FormattingReminder,message:b.L_CellFormatAmountBeyondLimits};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit]={name:b.L_MemoryLimit,message:b.L_CloseFileBeforeRetrieve};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile]={name:b.L_NetworkProblem,message:b.L_NetworkProblemRetrieveFile};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize]={name:b.L_InvalidValue,message:b.L_SliceSizeNotSupported};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened]={name:b.L_DisplayDialogError,message:b.L_DialogAlreadyOpened};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth]={name:b.L_IndexOutOfRange,message:b.L_IndexOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight]={name:b.L_IndexOutOfRange,message:b.L_IndexOutOfRange};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError]={name:b.L_DisplayDialogError,message:b.L_NetworkProblem};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme]={name:b.L_DialogNavigateError,message:b.L_DialogInvalidScheme};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains]={name:b.L_DisplayDialogError,message:b.L_DialogAddressNotTrusted};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS]={name:b.L_DisplayDialogError,message:b.L_DialogRequireHTTPS};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore]={name:b.L_DisplayDialogError,message:b.L_UserClickIgnore};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone]={name:b.L_DisplayDialogError,message:b.L_NewWindowCrossZoneErrorString};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave]={name:b.L_APINotSupported,message:b.L_InvalidSSOAddinMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn]={name:b.L_UserNotSignedIn,message:b.L_UserNotSignedIn};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted]={name:b.L_UserAborted,message:b.L_UserAbortedMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity]={name:b.L_UnsupportedUserIdentity,message:b.L_UnsupportedUserIdentityMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl]={name:b.L_InvalidResourceUrl,message:b.L_InvalidResourceUrlMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant]={name:b.L_InvalidGrant,message:b.L_InvalidGrantMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError]={name:b.L_SSOClientError,message:b.L_SSOClientErrorMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError]={name:b.L_SSOServerError,message:b.L_SSOServerErrorMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken]={name:b.L_AddinIsAlreadyRequestingToken,message:b.L_AddinIsAlreadyRequestingTokenMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory]={name:b.L_SSOUserConsentNotSupportedByCurrentAddinCategory,message:b.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage};a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost]={name:b.L_SSOConnectionLostError,message:b.L_SSOConnectionLostErrorMessage}}}}();var OfficeExt;(function(a){var b;(function(b){var a=1.1,z=function(){function a(){}return a}();b.RequirementVersion=z;var d=function(){function a(b){var a=this;a.isSetSupported=function(d,b){if(d==undefined)return false;if(b==undefined)b=0;var f=this._setMap,e=f._sets;if(e.hasOwnProperty(d.toLowerCase())){var g=e[d.toLowerCase()];try{var a=this._getVersion(g);b=b+"";var c=this._getVersion(b);if(a.major>0&&a.major>c.major)return true;if(a.minor>0&&a.minor>0&&a.major==c.major&&a.minor>=c.minor)return true}catch(h){return false}}return false};a._getVersion=function(e){var a="version format incorrect",b=e.split("."),c=0,d=0;if(b.length<2&&isNaN(Number(e)))throw a;else{c=Number(b[0]);if(b.length>=2)d=Number(b[1]);if(isNaN(c)||isNaN(d))throw a}var f={minor:d,major:c};return f};a._setMap=b;a.isSetSupported=a.isSetSupported.bind(a)}return a}();b.RequirementMatrix=d;var c=function(){function a(a){this._addSetMap=function(a){for(var b in a)this._sets[b]=a[b]};this._sets=a}return a}();b.DefaultSetRequirement=c;var x=function(c){__extends(b,c);function b(){c.call(this,{dialogapi:a})}return b}(c);b.DefaultDialogSetRequirement=x;var f=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,documentevents:a,excelapi:a,matrixbindings:a,matrixcoercion:a,selection:a,settings:a,tablebindings:a,tablecoercion:a,textbindings:a,textcoercion:a})}return b}(c);b.ExcelClientDefaultSetRequirement=f;var k=function(c){__extends(b,c);function b(){c.call(this);this._addSetMap({imagecoercion:a})}return b}(f);b.ExcelClientV1DefaultSetRequirement=k;var l=function(b){__extends(a,b);function a(){b.call(this,{mailbox:1.3})}return a}(c);b.OutlookClientDefaultSetRequirement=l;var h=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,compressedfile:a,customxmlparts:a,documentevents:a,file:a,htmlcoercion:a,matrixbindings:a,matrixcoercion:a,ooxmlcoercion:a,pdffile:a,selection:a,settings:a,tablebindings:a,tablecoercion:a,textbindings:a,textcoercion:a,textfile:a,wordapi:a})}return b}(c);b.WordClientDefaultSetRequirement=h;var p=function(c){__extends(b,c);function b(){c.call(this);this._addSetMap({customxmlparts:1.2,wordapi:1.2,imagecoercion:a})}return b}(h);b.WordClientV1DefaultSetRequirement=p;var e=function(c){__extends(b,c);function b(){c.call(this,{activeview:a,compressedfile:a,documentevents:a,file:a,pdffile:a,selection:a,settings:a,textcoercion:a})}return b}(c);b.PowerpointClientDefaultSetRequirement=e;var j=function(c){__extends(b,c);function b(){c.call(this);this._addSetMap({imagecoercion:a})}return b}(e);b.PowerpointClientV1DefaultSetRequirement=j;var o=function(c){__extends(b,c);function b(){c.call(this,{selection:a,textcoercion:a})}return b}(c);b.ProjectClientDefaultSetRequirement=o;var u=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,documentevents:a,matrixbindings:a,matrixcoercion:a,selection:a,settings:a,tablebindings:a,tablecoercion:a,textbindings:a,textcoercion:a,file:a})}return b}(c);b.ExcelWebDefaultSetRequirement=u;var w=function(c){__extends(b,c);function b(){c.call(this,{compressedfile:a,documentevents:a,file:a,imagecoercion:a,matrixcoercion:a,ooxmlcoercion:a,pdffile:a,selection:a,settings:a,tablecoercion:a,textcoercion:a,textfile:a})}return b}(c);b.WordWebDefaultSetRequirement=w;var n=function(c){__extends(b,c);function b(){c.call(this,{activeview:a,settings:a})}return b}(c);b.PowerpointWebDefaultSetRequirement=n;var g=function(b){__extends(a,b);function a(){b.call(this,{mailbox:1.3})}return a}(c);b.OutlookWebDefaultSetRequirement=g;var v=function(c){__extends(b,c);function b(){c.call(this,{activeview:a,documentevents:a,selection:a,settings:a,textcoercion:a})}return b}(c);b.SwayWebDefaultSetRequirement=v;var r=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,partialtablebindings:a,settings:a,tablebindings:a,tablecoercion:a})}return b}(c);b.AccessWebDefaultSetRequirement=r;var t=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,documentevents:a,matrixbindings:a,matrixcoercion:a,selection:a,settings:a,tablebindings:a,tablecoercion:a,textbindings:a,textcoercion:a})}return b}(c);b.ExcelIOSDefaultSetRequirement=t;var i=function(c){__extends(b,c);function b(){c.call(this,{bindingevents:a,compressedfile:a,customxmlparts:a,documentevents:a,file:a,htmlcoercion:a,matrixbindings:a,matrixcoercion:a,ooxmlcoercion:a,pdffile:a,selection:a,settings:a,tablebindings:a,tablecoercion:a,textbindings:a,textcoercion:a,textfile:a})}return b}(c);b.WordIOSDefaultSetRequirement=i;var s=function(b){__extends(a,b);function a(){b.call(this);this._addSetMap({customxmlparts:1.2,wordapi:1.2})}return a}(i);b.WordIOSV1DefaultSetRequirement=s;var m=function(c){__extends(b,c);function b(){c.call(this,{activeview:a,compressedfile:a,documentevents:a,file:a,pdffile:a,selection:a,settings:a,textcoercion:a})}return b}(c);b.PowerpointIOSDefaultSetRequirement=m;var q=function(c){__extends(b,c);function b(){c.call(this,{mailbox:a})}return b}(c);b.OutlookIOSDefaultSetRequirement=q;var y=function(){var b="undefined";function a(){}a.initializeOsfDda=function(){OSF.OUtil.setNamespace("Requirement",OSF.DDA)};a.getDefaultRequirementMatrix=function(f){this.initializeDefaultSetMatrix();var e=undefined,g=f.get_requirementMatrix();if(g!=undefined&&g.length>0&&typeof JSON!==b){var i=JSON.parse(f.get_requirementMatrix().toLowerCase());e=new d(new c(i))}else{var h=a.getClientFullVersionString(f);if(a.DefaultSetArrayMatrix!=undefined&&a.DefaultSetArrayMatrix[h]!=undefined)e=new d(a.DefaultSetArrayMatrix[h]);else e=new d(new c({}))}return e};a.getDefaultDialogRequirementMatrix=function(f){var a=undefined,e=f.get_dialogRequirementMatrix();if(e!=undefined&&e.length>0&&typeof JSON!==b){var g=JSON.parse(f.get_requirementMatrix().toLowerCase());a=new d(new c(g))}else a=new d(new x);return a};a.getClientFullVersionString=function(a){var d=a.get_appMinorVersion(),e="",b="",c=a.get_appName(),f=c==1024||c==4096||c==8192||c==65536;if(f&&a.get_appVersion()==1)if(c==4096&&d>=15)b="16.00.01";else b="16.00";else if(a.get_appName()==64)b=a.get_appVersion();else{if(d<10)e="0"+d;else e=""+d;b=a.get_appVersion()+"."+e}return a.get_appName()+"-"+b};a.initializeDefaultSetMatrix=function(){a.DefaultSetArrayMatrix[a.Excel_RCLIENT_1600]=new f;a.DefaultSetArrayMatrix[a.Word_RCLIENT_1600]=new h;a.DefaultSetArrayMatrix[a.PowerPoint_RCLIENT_1600]=new e;a.DefaultSetArrayMatrix[a.Excel_RCLIENT_1601]=new k;a.DefaultSetArrayMatrix[a.Word_RCLIENT_1601]=new p;a.DefaultSetArrayMatrix[a.PowerPoint_RCLIENT_1601]=new j;a.DefaultSetArrayMatrix[a.Outlook_RCLIENT_1600]=new l;a.DefaultSetArrayMatrix[a.Excel_WAC_1600]=new u;a.DefaultSetArrayMatrix[a.Word_WAC_1600]=new w;a.DefaultSetArrayMatrix[a.Outlook_WAC_1600]=new g;a.DefaultSetArrayMatrix[a.Outlook_WAC_1601]=new g;a.DefaultSetArrayMatrix[a.Project_RCLIENT_1600]=new o;a.DefaultSetArrayMatrix[a.Access_WAC_1600]=new r;a.DefaultSetArrayMatrix[a.PowerPoint_WAC_1600]=new n;a.DefaultSetArrayMatrix[a.Excel_IOS_1600]=new t;a.DefaultSetArrayMatrix[a.SWAY_WAC_1600]=new v;a.DefaultSetArrayMatrix[a.Word_IOS_1600]=new i;a.DefaultSetArrayMatrix[a.Word_IOS_16001]=new s;a.DefaultSetArrayMatrix[a.PowerPoint_IOS_1600]=new m;a.DefaultSetArrayMatrix[a.Outlook_IOS_1600]=new q};a.Excel_RCLIENT_1600="1-16.00";a.Excel_RCLIENT_1601="1-16.01";a.Word_RCLIENT_1600="2-16.00";a.Word_RCLIENT_1601="2-16.01";a.PowerPoint_RCLIENT_1600="4-16.00";a.PowerPoint_RCLIENT_1601="4-16.01";a.Outlook_RCLIENT_1600="8-16.00";a.Excel_WAC_1600="16-16.00";a.Word_WAC_1600="32-16.00";a.Outlook_WAC_1600="64-16.00";a.Outlook_WAC_1601="64-16.01";a.Project_RCLIENT_1600="128-16.00";a.Access_WAC_1600="256-16.00";a.PowerPoint_WAC_1600="512-16.00";a.Excel_IOS_1600="1024-16.00";a.SWAY_WAC_1600="2048-16.00";a.Word_IOS_1600="4096-16.00";a.Word_IOS_16001="4096-16.00.01";a.PowerPoint_IOS_1600="8192-16.00";a.Outlook_IOS_1600="65536-16.00";a.DefaultSetArrayMatrix={};return a}();b.RequirementsMatrixFactory=y})(b=a.Requirement||(a.Requirement={}))})(OfficeExt||(OfficeExt={}));OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();var OfficeExt;(function(a){var b;(function(a){var b=function(){function a(){var a=this;a.getDiagnostics=function(b){var a={host:this.getHost(),version:b||this.getDefaultVersion(),platform:this.getPlatform()};return a};a.platformRemappings={web:Microsoft.Office.WebExtension.PlatformType.OfficeOnline,winrt:Microsoft.Office.WebExtension.PlatformType.Universal,win32:Microsoft.Office.WebExtension.PlatformType.PC,mac:Microsoft.Office.WebExtension.PlatformType.Mac,ios:Microsoft.Office.WebExtension.PlatformType.iOS,android:Microsoft.Office.WebExtension.PlatformType.Android};a.camelCaseMappings={powerpoint:Microsoft.Office.WebExtension.HostType.PowerPoint,onenote:Microsoft.Office.WebExtension.HostType.OneNote};a.hostInfo=OSF._OfficeAppFactory.getHostInfo();a.getHost=a.getHost.bind(a);a.getPlatform=a.getPlatform.bind(a);a.getDiagnostics=a.getDiagnostics.bind(a)}a.prototype.capitalizeFirstLetter=function(a){if(a)return a[0].toUpperCase()+a.slice(1).toLowerCase();return a};a.getInstance=function(){if(a.hostObj===undefined)a.hostObj=new a;return a.hostObj};a.prototype.getPlatform=function(){var a=this;if(a.hostInfo.hostPlatform){var b=a.hostInfo.hostPlatform.toLowerCase();if(a.platformRemappings[b])return a.platformRemappings[b]}return null};a.prototype.getHost=function(){var a=this;if(a.hostInfo.hostType){var b=a.hostInfo.hostType.toLowerCase();if(a.camelCaseMappings[b])return a.camelCaseMappings[b];b=a.capitalizeFirstLetter(a.hostInfo.hostType);if(Microsoft.Office.WebExtension.HostType[b])return Microsoft.Office.WebExtension.HostType[b]}return null};a.prototype.getDefaultVersion=function(){if(this.getHost())return "16.0.0000.0000";return null};return a}();a.Host=b})(b=a.HostName||(a.HostName={}))})(OfficeExt||(OfficeExt={}));Microsoft.Office.WebExtension.ApplicationMode={WebEditor:"webEditor",WebViewer:"webViewer",Client:"client"};Microsoft.Office.WebExtension.DocumentMode={ReadOnly:"readOnly",ReadWrite:"readWrite"};OSF.NamespaceManager=function(){var b,a=false;return {enableShortcut:function(){if(!a){if(window.Office)b=window.Office;else OSF.OUtil.setNamespace("Office",window);window.Office=Microsoft.Office.WebExtension;a=true}},disableShortcut:function(){if(a){if(b)window.Office=b;else OSF.OUtil.unsetNamespace("Office",window);a=false}}}}();OSF.NamespaceManager.enableShortcut();Microsoft.Office.WebExtension.useShortNamespace=function(a){if(a)OSF.NamespaceManager.enableShortcut();else OSF.NamespaceManager.disableShortcut()};Microsoft.Office.WebExtension.select=function(a,b){var c;if(a&&typeof a=="string"){var d=a.indexOf("#");if(d!=-1){var h=a.substring(0,d),g=a.substring(d+1);switch(h){case "binding":case "bindings":if(g)c=new OSF.DDA.BindingPromise(g)}}}if(!c){if(b){var e=typeof b;if(e=="function"){var f={};f[Microsoft.Office.WebExtension.Parameters.Callback]=b;OSF.DDA.issueAsyncResult(f,OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext,OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext))}else throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction,e)}}else{c.onFail=b;return c}};OSF.DDA.Context=function(a,g,h,c,d){var f="requirements",b=this;OSF.OUtil.defineEnumerableProperties(b,{contentLanguage:{value:a.get_dataLocale()},displayLanguage:{value:a.get_appUILocale()},touchEnabled:{value:a.get_touchEnabled()},commerceAllowed:{value:a.get_commerceAllowed()},host:{value:OfficeExt.HostName.Host.getInstance().getHost()},platform:{value:OfficeExt.HostName.Host.getInstance().getPlatform()},diagnostics:{value:OfficeExt.HostName.Host.getInstance().getDiagnostics(a.get_hostFullVersion())}});h&&OSF.OUtil.defineEnumerableProperty(b,"license",{value:h});a.ui&&OSF.OUtil.defineEnumerableProperty(b,"ui",{value:a.ui});a.auth&&OSF.OUtil.defineEnumerableProperty(b,"auth",{value:a.auth});a.application&&OSF.OUtil.defineEnumerableProperty(b,"application",{value:a.application});if(a.get_isDialog()){var e=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(a);OSF.OUtil.defineEnumerableProperty(b,f,{value:e})}else{g&&OSF.OUtil.defineEnumerableProperty(b,"document",{value:g});if(c){var i=c.displayName||"appOM";delete c.displayName;OSF.OUtil.defineEnumerableProperty(b,i,{value:c})}d&&OSF.OUtil.defineEnumerableProperty(b,"officeTheme",{"get":function(){return d()}});var e=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(a);OSF.OUtil.defineEnumerableProperty(b,f,{value:e})}};OSF.DDA.OutlookContext=function(c,a,d,e,b){OSF.DDA.OutlookContext.uber.constructor.call(this,c,null,d,e,b);a&&OSF.OUtil.defineEnumerableProperty(this,"roamingSettings",{value:a})};OSF.OUtil.extend(OSF.DDA.OutlookContext,OSF.DDA.Context);OSF.DDA.OutlookAppOm=function(){};OSF.DDA.Application=function(){};OSF.DDA.Document=function(b,c){var a;switch(b.get_clientMode()){case OSF.ClientMode.ReadOnly:a=Microsoft.Office.WebExtension.DocumentMode.ReadOnly;break;case OSF.ClientMode.ReadWrite:a=Microsoft.Office.WebExtension.DocumentMode.ReadWrite}c&&OSF.OUtil.defineEnumerableProperty(this,"settings",{value:c});OSF.OUtil.defineMutableProperties(this,{mode:{value:a},url:{value:b.get_docUrl()}})};OSF.DDA.JsomDocument=function(d,b,e){var a=this;OSF.DDA.JsomDocument.uber.constructor.call(a,d,e);b&&OSF.OUtil.defineEnumerableProperty(a,"bindings",{"get":function(){return b}});var c=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(a,[c.GetSelectedDataAsync,c.SetSelectedDataAsync]);OSF.DDA.DispIdHost.addEventSupport(a,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))};OSF.OUtil.extend(OSF.DDA.JsomDocument,OSF.DDA.Document);OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension,"context",{"get":function(){var a;if(OSF&&OSF._OfficeAppFactory)a=OSF._OfficeAppFactory.getContext();return a}});OSF.DDA.License=function(a){OSF.OUtil.defineEnumerableProperty(this,"value",{value:a})};OSF.DDA.ApiMethodCall=function(c,f,e,g,h){var a=this,d=c.length,b=OSF.OUtil.delayExecutionAndCache(function(){return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters,h)});a.verifyArguments=function(d,f){for(var e in d){var a=d[e],c=f[e];if(a["enum"])switch(typeof c){case "string":if(OSF.OUtil.listContainsValue(a["enum"],c))break;case "undefined":throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;default:throw b()}if(a["types"])if(!OSF.OUtil.listContainsValue(a["types"],typeof c))throw b()}};a.extractRequiredArguments=function(g,l,j){if(g.length<d)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);for(var e=[],a=0;a<d;a++)e.push(g[a]);this.verifyArguments(c,e);var i={};for(a=0;a<d;a++){var f=c[a],h=e[a];if(f.verify){var k=f.verify(h,l,j);if(!k)throw b()}i[f.name]=h}return i},a.fillOptions=function(a,e,h,g){a=a||{};for(var d in f)if(!OSF.OUtil.listContainsKey(a,d)){var c=undefined,b=f[d];if(b.calculate&&e)c=b.calculate(e,h,g);if(!c&&b.defaultValue!==undefined)c=b.defaultValue;a[d]=c}return a};a.constructCallArgs=function(c,d,f,b){var a={};for(var i in c)a[i]=c[i];for(var h in d)a[h]=d[h];for(var j in e)a[j]=e[j](f,b);if(g)a=g(a,f,b);return a}};OSF.OUtil.setNamespace("AsyncResultEnum",OSF.DDA);OSF.DDA.AsyncResultEnum.Properties={Context:"Context",Value:"Value",Status:"Status",Error:"Error"};Microsoft.Office.WebExtension.AsyncResultStatus={Succeeded:"succeeded",Failed:"failed"};OSF.DDA.AsyncResultEnum.ErrorCode={Success:0,Failed:1};OSF.DDA.AsyncResultEnum.ErrorProperties={Name:"Name",Message:"Message",Code:"Code"};OSF.DDA.AsyncMethodNames={};OSF.DDA.AsyncMethodNames.addNames=function(b){for(var a in b){var c={};OSF.OUtil.defineEnumerableProperties(c,{id:{value:a},displayName:{value:b[a]}});OSF.DDA.AsyncMethodNames[a]=c}};OSF.DDA.AsyncMethodCall=function(d,e,i,f,g,j,k){var a="function",c=d.length,b=new OSF.DDA.ApiMethodCall(d,e,i,j,k);function h(h,j,l,k){if(h.length>c+2)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);for(var d,f,i=h.length-1;i>=c;i--){var g=h[i];switch(typeof g){case "object":if(d)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);else d=g;break;case a:if(f)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);else f=g;break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument)}}d=b.fillOptions(d,j,l,k);if(f)if(d[Microsoft.Office.WebExtension.Parameters.Callback])throw Strings.OfficeOM.L_RedundantCallbackSpecification;else d[Microsoft.Office.WebExtension.Parameters.Callback]=f;b.verifyArguments(e,d);return d}this.verifyAndExtractCall=function(e,c,a){var d=b.extractRequiredArguments(e,c,a),g=h(e,d,c,a),f=b.constructCallArgs(d,g,c,a);return f};this.processResponse=function(c,b,e,d){var a;if(c==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)if(f)a=f(b,e,d);else a=b;else if(g)a=g(c,b);else a=OSF.DDA.ErrorCodeManager.getErrorArgs(c);return a};this.getCallArgs=function(g){for(var b,d,f=g.length-1;f>=c;f--){var e=g[f];switch(typeof e){case "object":b=e;break;case a:d=e}}b=b||{};if(d)b[Microsoft.Office.WebExtension.Parameters.Callback]=d;return b}};OSF.DDA.AsyncMethodCallFactory=function(){return {manufacture:function(a){var c=a.supportedOptions?OSF.OUtil.createObject(a.supportedOptions):[],b=a.privateStateCallbacks?OSF.OUtil.createObject(a.privateStateCallbacks):[];return new OSF.DDA.AsyncMethodCall(a.requiredArguments||[],c,b,a.onSucceeded,a.onFailed,a.checkCallArgs,a.method.displayName)}}}();OSF.DDA.AsyncMethodCalls={};OSF.DDA.AsyncMethodCalls.define=function(a){OSF.DDA.AsyncMethodCalls[a.method.id]=OSF.DDA.AsyncMethodCallFactory.manufacture(a)};OSF.DDA.Error=function(c,a,b){OSF.OUtil.defineEnumerableProperties(this,{name:{value:c},message:{value:a},code:{value:b}})};OSF.DDA.AsyncResult=function(b,a){OSF.OUtil.defineEnumerableProperties(this,{value:{value:b[OSF.DDA.AsyncResultEnum.Properties.Value]},status:{value:a?Microsoft.Office.WebExtension.AsyncResultStatus.Failed:Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded}});b[OSF.DDA.AsyncResultEnum.Properties.Context]&&OSF.OUtil.defineEnumerableProperty(this,"asyncContext",{value:b[OSF.DDA.AsyncResultEnum.Properties.Context]});a&&OSF.OUtil.defineEnumerableProperty(this,"error",{value:new OSF.DDA.Error(a[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],a[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],a[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])})};OSF.DDA.issueAsyncResult=function(d,f,a){var e=d[Microsoft.Office.WebExtension.Parameters.Callback];if(e){var c={};c[OSF.DDA.AsyncResultEnum.Properties.Context]=d[Microsoft.Office.WebExtension.Parameters.AsyncContext];var b;if(f==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)c[OSF.DDA.AsyncResultEnum.Properties.Value]=a;else{b={};a=a||OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);b[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=f||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;b[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=a.name||a;b[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=a.message||a}e(new OSF.DDA.AsyncResult(c,b))}};OSF.DDA.SyncMethodNames={};OSF.DDA.SyncMethodNames.addNames=function(b){for(var a in b){var c={};OSF.OUtil.defineEnumerableProperties(c,{id:{value:a},displayName:{value:b[a]}});OSF.DDA.SyncMethodNames[a]=c}};OSF.DDA.SyncMethodCall=function(b,c,f,g,h){var d=b.length,a=new OSF.DDA.ApiMethodCall(b,c,f,g,h);function e(e,h,j,i){if(e.length>d+1)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);for(var b,k,f=e.length-1;f>=d;f--){var g=e[f];switch(typeof g){case "object":if(b)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);else b=g;break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument)}}b=a.fillOptions(b,h,j,i);a.verifyArguments(c,b);return b}this.verifyAndExtractCall=function(f,c,b){var d=a.extractRequiredArguments(f,c,b),h=e(f,d,c,b),g=a.constructCallArgs(d,h,c,b);return g}};OSF.DDA.SyncMethodCallFactory=function(){return {manufacture:function(a){var b=a.supportedOptions?OSF.OUtil.createObject(a.supportedOptions):[];return new OSF.DDA.SyncMethodCall(a.requiredArguments||[],b,a.privateStateCallbacks,a.checkCallArgs,a.method.displayName)}}}();OSF.DDA.SyncMethodCalls={};OSF.DDA.SyncMethodCalls.define=function(a){OSF.DDA.SyncMethodCalls[a.method.id]=OSF.DDA.SyncMethodCallFactory.manufacture(a)};OSF.DDA.ListType=function(){var a={};return {setListType:function(c,b){a[c]=b},isListType:function(b){return OSF.OUtil.listContainsKey(a,b)},getDescriptor:function(b){return a[b]}}}();OSF.DDA.HostParameterMap=function(b,c){var j="fromHost",a=this,i="toHost",e=j,l="sourceData",g="self",d={};d[Microsoft.Office.WebExtension.Parameters.Data]={toHost:function(a){if(a!=null&&a.rows!==undefined){var b={};b[OSF.DDA.TableDataProperties.TableRows]=a.rows;b[OSF.DDA.TableDataProperties.TableHeaders]=a.headers;a=b}return a},fromHost:function(a){return a}};d[Microsoft.Office.WebExtension.Parameters.SampleData]=d[Microsoft.Office.WebExtension.Parameters.Data];function f(j,i){var m=j?{}:undefined;for(var h in j){var g=j[h],a;if(OSF.DDA.ListType.isListType(h)){a=[];for(var n in g)a.push(f(g[n],i))}else if(OSF.OUtil.listContainsKey(d,h))a=d[h][i](g);else if(i==e&&b.preserveNesting(h))a=f(g,i);else{var k=c[h];if(k){var l=k[i];if(l){a=l[g];if(a===undefined)a=g}}else a=g}m[h]=a}return m}function k(j,h){var e;for(var a in h){var d;if(b.isComplexType(a))d=k(j,c[a][i]);else d=j[a];if(d!=undefined){if(!e)e={};var f=h[a];if(f==g)f=a;e[f]=b.pack(a,d)}}return e}function h(j,n,f){if(!f)f={};for(var a in n){var k=n[a],d;if(k==g)d=j;else if(k==l){f[a]=j.toArray();continue}else d=j[k];if(d===null||d===undefined)f[a]=undefined;else{d=b.unpack(a,d);var i;if(b.isComplexType(a)){i=c[a][e];if(b.preserveNesting(a))f[a]=h(d,i);else h(d,i,f)}else if(OSF.DDA.ListType.isListType(a)){i={};var p=OSF.DDA.ListType.getDescriptor(a);i[p]=g;var m=new Array(d.length);for(var o in d)m[o]=h(d[o],i);f[a]=m}else f[a]=d}}return f}function m(l,e,a){var d=c[l][a],b;if(a=="toHost"){var i=f(e,a);b=k(i,d)}else if(a==j){var g=h(e,d);b=f(g,a)}return b}if(!c)c={};a.addMapping=function(l,h){var a,d;if(h.map){a=h.map;d={};for(var j in a){var k=a[j];if(k==g)k=j;d[k]=j}}else{a=h.toHost;d=h.fromHost}var b=c[l];if(b){var f=b[i];for(var n in f)a[n]=f[n];f=b[e];for(var m in f)d[m]=f[m]}else b=c[l]={};b[i]=a;b[e]=d};a.toHost=function(b,a){return m(b,a,i)};a.fromHost=function(a,b){return m(a,b,e)};a.self=g;a.sourceData=l;a.addComplexType=function(a){b.addComplexType(a)};a.getDynamicType=function(a){return b.getDynamicType(a)};a.setDynamicType=function(c,a){b.setDynamicType(c,a)};a.dynamicTypes=d;a.doMapValues=function(a,b){return f(a,b)}};OSF.DDA.SpecialProcessor=function(c,b){var a=this;a.addComplexType=function(a){c.push(a)};a.getDynamicType=function(a){return b[a]};a.setDynamicType=function(c,a){b[c]=a};a.isComplexType=function(a){return OSF.OUtil.listContainsValue(c,a)};a.isDynamicType=function(a){return OSF.OUtil.listContainsKey(b,a)};a.preserveNesting=function(b){var a=[];OSF.DDA.PropertyDescriptors&&a.push(OSF.DDA.PropertyDescriptors.Subset);if(OSF.DDA.DataNodeEventProperties)a=a.concat([OSF.DDA.DataNodeEventProperties.OldNode,OSF.DDA.DataNodeEventProperties.NewNode,OSF.DDA.DataNodeEventProperties.NextSiblingNode]);return OSF.OUtil.listContainsValue(a,b)};a.pack=function(c,d){var a;if(this.isDynamicType(c))a=b[c].toHost(d);else a=d;return a};a.unpack=function(c,d){var a;if(this.isDynamicType(c))a=b[c].fromHost(d);else a=d;return a}};OSF.DDA.getDecoratedParameterMap=function(d,c){var a=new OSF.DDA.HostParameterMap(d),f=a.self;function b(a){var c=null;if(a){c={};for(var d=a.length,b=0;b<d;b++)c[a[b].name]=a[b].value}return c}a.define=function(c){var d={},e=b(c.toHost);if(c.invertible)d.map=e;else if(c.canonical)d.toHost=d.fromHost=e;else{d.toHost=e;d.fromHost=b(c.fromHost)}a.addMapping(c.type,d);c.isComplexType&&a.addComplexType(c.type)};for(var e in c)a.define(c[e]);return a};OSF.OUtil.setNamespace("DispIdHost",OSF.DDA);OSF.DDA.DispIdHost.Methods={InvokeMethod:"invokeMethod",AddEventHandler:"addEventHandler",RemoveEventHandler:"removeEventHandler",OpenDialog:"openDialog",CloseDialog:"closeDialog",MessageParent:"messageParent",SendMessage:"sendMessage"};OSF.DDA.DispIdHost.Delegates={ExecuteAsync:"executeAsync",RegisterEventAsync:"registerEventAsync",UnregisterEventAsync:"unregisterEventAsync",ParameterMap:"parameterMap",OpenDialog:"openDialog",CloseDialog:"closeDialog",MessageParent:"messageParent",SendMessage:"sendMessage"};OSF.DDA.DispIdHost.Facade=function(f,h){var b=false,d=null,g=this,c={},e=OSF.DDA.AsyncMethodNames,a=OSF.DDA.MethodDispId,n={GoToByIdAsync:a.dispidNavigateToMethod,GetSelectedDataAsync:a.dispidGetSelectedDataMethod,SetSelectedDataAsync:a.dispidSetSelectedDataMethod,GetDocumentCopyChunkAsync:a.dispidGetDocumentCopyChunkMethod,ReleaseDocumentCopyAsync:a.dispidReleaseDocumentCopyMethod,GetDocumentCopyAsync:a.dispidGetDocumentCopyMethod,AddFromSelectionAsync:a.dispidAddBindingFromSelectionMethod,AddFromPromptAsync:a.dispidAddBindingFromPromptMethod,AddFromNamedItemAsync:a.dispidAddBindingFromNamedItemMethod,GetAllAsync:a.dispidGetAllBindingsMethod,GetByIdAsync:a.dispidGetBindingMethod,ReleaseByIdAsync:a.dispidReleaseBindingMethod,GetDataAsync:a.dispidGetBindingDataMethod,SetDataAsync:a.dispidSetBindingDataMethod,AddRowsAsync:a.dispidAddRowsMethod,AddColumnsAsync:a.dispidAddColumnsMethod,DeleteAllDataValuesAsync:a.dispidClearAllRowsMethod,RefreshAsync:a.dispidLoadSettingsMethod,SaveAsync:a.dispidSaveSettingsMethod,GetActiveViewAsync:a.dispidGetActiveViewMethod,GetFilePropertiesAsync:a.dispidGetFilePropertiesMethod,GetOfficeThemeAsync:a.dispidGetOfficeThemeMethod,GetDocumentThemeAsync:a.dispidGetDocumentThemeMethod,ClearFormatsAsync:a.dispidClearFormatsMethod,SetTableOptionsAsync:a.dispidSetTableOptionsMethod,SetFormatsAsync:a.dispidSetFormatsMethod,GetAccessTokenAsync:a.dispidGetAccessTokenMethod,ExecuteRichApiRequestAsync:a.dispidExecuteRichApiRequestMethod,AppCommandInvocationCompletedAsync:a.dispidAppCommandInvocationCompletedMethod,CloseContainerAsync:a.dispidCloseContainerMethod,OpenBrowserWindow:a.dispidOpenBrowserWindow,CreateDocumentAsync:a.dispidCreateDocumentMethod,InsertFormAsync:a.dispidInsertFormMethod,AddDataPartAsync:a.dispidAddDataPartMethod,GetDataPartByIdAsync:a.dispidGetDataPartByIdMethod,GetDataPartsByNameSpaceAsync:a.dispidGetDataPartsByNamespaceMethod,GetPartXmlAsync:a.dispidGetDataPartXmlMethod,GetPartNodesAsync:a.dispidGetDataPartNodesMethod,DeleteDataPartAsync:a.dispidDeleteDataPartMethod,GetNodeValueAsync:a.dispidGetDataNodeValueMethod,GetNodeXmlAsync:a.dispidGetDataNodeXmlMethod,GetRelativeNodesAsync:a.dispidGetDataNodesMethod,SetNodeValueAsync:a.dispidSetDataNodeValueMethod,SetNodeXmlAsync:a.dispidSetDataNodeXmlMethod,AddDataPartNamespaceAsync:a.dispidAddDataNamespaceMethod,GetDataPartNamespaceAsync:a.dispidGetDataUriByPrefixMethod,GetDataPartPrefixAsync:a.dispidGetDataPrefixByUriMethod,GetNodeTextAsync:a.dispidGetDataNodeTextMethod,SetNodeTextAsync:a.dispidSetDataNodeTextMethod,GetSelectedTask:a.dispidGetSelectedTaskMethod,GetTask:a.dispidGetTaskMethod,GetWSSUrl:a.dispidGetWSSUrlMethod,GetTaskField:a.dispidGetTaskFieldMethod,GetSelectedResource:a.dispidGetSelectedResourceMethod,GetResourceField:a.dispidGetResourceFieldMethod,GetProjectField:a.dispidGetProjectFieldMethod,GetSelectedView:a.dispidGetSelectedViewMethod,GetTaskByIndex:a.dispidGetTaskByIndexMethod,GetResourceByIndex:a.dispidGetResourceByIndexMethod,SetTaskField:a.dispidSetTaskFieldMethod,SetResourceField:a.dispidSetResourceFieldMethod,GetMaxTaskIndex:a.dispidGetMaxTaskIndexMethod,GetMaxResourceIndex:a.dispidGetMaxResourceIndexMethod,CreateTask:a.dispidCreateTaskMethod};for(var i in n)if(e[i])c[e[i].id]=n[i];e=OSF.DDA.SyncMethodNames;a=OSF.DDA.MethodDispId;var m={MessageParent:a.dispidMessageParentMethod,SendMessage:a.dispidSendMessageMethod};for(var i in m)if(e[i])c[e[i].id]=m[i];e=Microsoft.Office.WebExtension.EventType;a=OSF.DDA.EventDispId;var o={SettingsChanged:a.dispidSettingsChangedEvent,DocumentSelectionChanged:a.dispidDocumentSelectionChangedEvent,BindingSelectionChanged:a.dispidBindingSelectionChangedEvent,BindingDataChanged:a.dispidBindingDataChangedEvent,ActiveViewChanged:a.dispidActiveViewChangedEvent,OfficeThemeChanged:a.dispidOfficeThemeChangedEvent,DocumentThemeChanged:a.dispidDocumentThemeChangedEvent,AppCommandInvoked:a.dispidAppCommandInvokedEvent,DialogMessageReceived:a.dispidDialogMessageReceivedEvent,DialogParentMessageReceived:a.dispidDialogParentMessageReceivedEvent,ObjectDeleted:a.dispidObjectDeletedEvent,ObjectSelectionChanged:a.dispidObjectSelectionChangedEvent,ObjectDataChanged:a.dispidObjectDataChangedEvent,ContentControlAdded:a.dispidContentControlAddedEvent,RichApiMessage:a.dispidRichApiMessageEvent,ItemChanged:a.dispidOlkItemSelectedChangedEvent,RecipientsChanged:a.dispidOlkRecipientsChangedEvent,AppointmentTimeChanged:a.dispidOlkAppointmentTimeChangedEvent,TaskSelectionChanged:a.dispidTaskSelectionChangedEvent,ResourceSelectionChanged:a.dispidResourceSelectionChangedEvent,ViewSelectionChanged:a.dispidViewSelectionChangedEvent,DataNodeInserted:a.dispidDataNodeAddedEvent,DataNodeReplaced:a.dispidDataNodeReplacedEvent,DataNodeDeleted:a.dispidDataNodeDeletedEvent};for(var k in o)if(e[k])c[e[k]]=o[k];function l(a){return a==OSF.DDA.EventDispId.dispidObjectDeletedEvent||a==OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent||a==OSF.DDA.EventDispId.dispidObjectDataChangedEvent||a==OSF.DDA.EventDispId.dispidContentControlAddedEvent}function j(a,c,d,b){if(typeof a=="number"){if(!b)b=c.getCallArgs(d);OSF.DDA.issueAsyncResult(b,a,OSF.DDA.ErrorCodeManager.getErrorArgs(a))}else throw a}g[OSF.DDA.DispIdHost.Methods.InvokeMethod]=function(t,m,n,q){var a;try{var i=t.id,l=OSF.DDA.AsyncMethodCalls[i];a=l.verifyAndExtractCall(m,n,q);var k=c[i],s=f(i),b=d;if(window.Excel&&window.Office.context.requirements.isSetSupported("RedirectV1Api"))window.Excel._RedirectV1APIs=true;if(window.Excel&&window.Excel._RedirectV1APIs&&(b=window.Excel._V1APIMap[i])){var e=OSF.OUtil.shallowCopy(a);delete e[Microsoft.Office.WebExtension.Parameters.AsyncContext];if(b.preprocess)e=b.preprocess(e);var o=new window.Excel.RequestContext,u=b.call(o,e);o.sync().then(function(){var c=u.value,d=c.status;delete c["status"];delete c["@odata.type"];if(b.postprocess)c=b.postprocess(c,e);if(d!=0)c=OSF.DDA.ErrorCodeManager.getErrorArgs(d);OSF.DDA.issueAsyncResult(a,d,c)})["catch"](function(){OSF.DDA.issueAsyncResult(a,OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure,d)})}else{var g;if(h.toHost)g=h.toHost(k,a);else g=a;var r=(new Date).getTime();s[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({dispId:k,hostCallArgs:g,onCalling:function(){},onReceiving:function(){},onComplete:function(c,d){var b;if(c==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)if(h.fromHost)b=h.fromHost(k,d);else b=d;else b=d;var e=l.processResponse(c,b,n,a);OSF.DDA.issueAsyncResult(a,c,e);OSF.AppTelemetry&&OSF.AppTelemetry.onMethodDone(k,g,Math.abs((new Date).getTime()-r),c)}})}}catch(p){j(p,l,m,a)}};g[OSF.DDA.DispIdHost.Methods.AddEventHandler]=function(p,d,n,s){var e,a,m,g=b;function k(b){if(b==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess){var f=!g?d.addEventHandler(a,m):d.addObjectEventHandler(a,e[Microsoft.Office.WebExtension.Parameters.Id],m);if(!f)b=OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed}var c;if(b!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)c=OSF.DDA.ErrorCodeManager.getErrorArgs(b);OSF.DDA.issueAsyncResult(e,b,c)}try{var q=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];e=q.verifyAndExtractCall(p,n,d);a=e[Microsoft.Office.WebExtension.Parameters.EventType];m=e[Microsoft.Office.WebExtension.Parameters.Handler];if(s){k(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);return}var o=c[a];g=l(o);var i=g?e[Microsoft.Office.WebExtension.Parameters.Id]:n.id||"",u=g?d.getObjectEventHandlerCount(a,i):d.getEventHandlerCount(a);if(u==0){var t=f(a)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];t({eventType:a,dispId:o,targetId:i,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:k,onEvent:function(c){var b=h.fromHost(o,c);if(!g)d.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(a,n,b));else d.fireObjectEvent(i,OSF.DDA.OMFactory.manufactureEventArgs(a,i,b))}})}else k(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)}catch(r){j(r,q,p,e)}};g[OSF.DDA.DispIdHost.Methods.RemoveEventHandler]=function(p,e,r){var g,a,m,h=b;function o(a){var b;if(a!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)b=OSF.DDA.ErrorCodeManager.getErrorArgs(a);OSF.DDA.issueAsyncResult(g,a,b)}try{var q=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];g=q.verifyAndExtractCall(p,r,e);a=g[Microsoft.Office.WebExtension.Parameters.EventType];m=g[Microsoft.Office.WebExtension.Parameters.Handler];var s=c[a];h=l(s);var k=h?g[Microsoft.Office.WebExtension.Parameters.Id]:r.id||"",n,i;if(m===d){i=h?e.clearObjectEventHandlers(a,k):e.clearEventHandlers(a);n=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess}else{i=h?e.removeObjectEventHandler(a,k,m):e.removeEventHandler(a,m);n=i?OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess:OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist}var v=h?e.getObjectEventHandlerCount(a,k):e.getEventHandlerCount(a);if(i&&v==0){var u=f(a)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];u({eventType:a,dispId:s,targetId:k,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:o})}else o(n)}catch(t){j(t,q,p,g)}};g[OSF.DDA.DispIdHost.Methods.OpenDialog]=function(p,a,o){var i,n,e=Microsoft.Office.WebExtension.EventType.DialogMessageReceived,g=Microsoft.Office.WebExtension.EventType.DialogEventReceived;function k(b){var d;if(b!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)d=OSF.DDA.ErrorCodeManager.getErrorArgs(b);else{var c={};c[Microsoft.Office.WebExtension.Parameters.Id]=n;c[Microsoft.Office.WebExtension.Parameters.Data]=a;var d=l.processResponse(b,c,o,i);OSF.DialogShownStatus.hasDialogShown=true;a.clearEventHandlers(e);a.clearEventHandlers(g)}OSF.DDA.issueAsyncResult(i,b,d)}try{(e==undefined||g==undefined)&&k(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);if(OSF.DDA.AsyncMethodNames.DisplayDialogAsync==d){k(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);return}var l=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];i=l.verifyAndExtractCall(p,o,a);var q=c[e],m=f(e),s=m[OSF.DDA.DispIdHost.Delegates.OpenDialog]!=undefined?m[OSF.DDA.DispIdHost.Delegates.OpenDialog]:m[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];n=JSON.stringify(i);if(!OSF.DialogShownStatus.hasDialogShown){a.clearQueuedEvent(e);a.clearQueuedEvent(g);a.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived)}s({eventType:e,dispId:q,targetId:n,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:k,onEvent:function(j){var i=h.fromHost(q,j),f=OSF.DDA.OMFactory.manufactureEventArgs(e,o,i);if(f.type==g){var d=OSF.DDA.ErrorCodeManager.getErrorArgs(f.error),c={};c[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;c[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=d.name||d;c[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=d.message||d;f.error=new OSF.DDA.Error(c[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],c[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],c[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])}a.fireOrQueueEvent(f);if(i[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogClosed){a.clearEventHandlers(e);a.clearEventHandlers(g);a.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);OSF.DialogShownStatus.hasDialogShown=b}}})}catch(r){j(r,l,p,i)}};g[OSF.DDA.DispIdHost.Methods.CloseDialog]=function(h,o,e,q){var l,a,i,g=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;function n(a){g=a;OSF.DialogShownStatus.hasDialogShown=b}try{var k=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];l=k.verifyAndExtractCall(h,q,e);a=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;i=Microsoft.Office.WebExtension.EventType.DialogEventReceived;e.clearEventHandlers(a);e.clearEventHandlers(i);var r=c[a],d=f(a),p=d[OSF.DDA.DispIdHost.Delegates.CloseDialog]!=undefined?d[OSF.DDA.DispIdHost.Delegates.CloseDialog]:d[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];p({eventType:a,dispId:r,targetId:o,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:n})}catch(m){j(m,k,h,l)}if(g!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed,OSF.DDA.AsyncMethodNames.CloseAsync.displayName,g)};g[OSF.DDA.DispIdHost.Methods.MessageParent]=function(a,i){var d={},b=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id],e=b.verifyAndExtractCall(a,i,d),g=f(OSF.DDA.SyncMethodNames.MessageParent.id),h=g[OSF.DDA.DispIdHost.Delegates.MessageParent],j=c[OSF.DDA.SyncMethodNames.MessageParent.id];return h({dispId:j,hostCallArgs:e,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)}})};g[OSF.DDA.DispIdHost.Methods.SendMessage]=function(a,k,i){var d={},b=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id],e=b.verifyAndExtractCall(a,i,d),g=f(OSF.DDA.SyncMethodNames.SendMessage.id),h=g[OSF.DDA.DispIdHost.Delegates.SendMessage],j=c[OSF.DDA.SyncMethodNames.SendMessage.id];return h({dispId:j,hostCallArgs:e,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)}})}};OSF.DDA.DispIdHost.addAsyncMethods=function(a,b,e){for(var f in b){var c=b[f],d=c.displayName;!a[d]&&OSF.OUtil.defineEnumerableProperty(a,d,{value:function(b){return function(){var c=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];c(b,arguments,a,e)}}(c)})}};OSF.DDA.DispIdHost.addEventSupport=function(a,b,e){var d=OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName,c=OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;!a[d]&&OSF.OUtil.defineEnumerableProperty(a,d,{value:function(){var c=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];c(arguments,b,a,e)}});!a[c]&&OSF.OUtil.defineEnumerableProperty(a,c,{value:function(){var c=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];c(arguments,b,a)}})};var OfficeExt;(function(g){var c="\n",d=true,a=null,b="undefined",j=function(){function c(){}c.isInstanceOfType=function(f,e){if(typeof e===b||e===a)return false;if(e instanceof f)return d;var c=e.constructor;if(!c||typeof c!=="function"||!c.__typeName||c.__typeName==="Object")c=Object;return !!(c===f)||c.__typeName&&f.__typeName&&c.__typeName===f.__typeName};return c}();g.MsAjaxTypeHelper=j;var h=function(){var e="Parameter name: {0}";function d(){}d.create=function(c,b){var a=new Error(c);a.message=c;if(b)for(var d in b)a[d]=b[d];a.popStackFrame();return a};d.parameterCount=function(a){var c="Sys.ParameterCountException: "+(a?a:"Parameter count mismatch."),b=d.create(c,{name:"Sys.ParameterCountException"});b.popStackFrame();return b};d.argument=function(a,g){var b="Sys.ArgumentException: "+(g?g:"Value does not fall within the expected range.");if(a)b+=c+f.format(e,a);var h=d.create(b,{name:"Sys.ArgumentException",paramName:a});h.popStackFrame();return h};d.argumentNull=function(a,g){var b="Sys.ArgumentNullException: "+(g?g:"Value cannot be null.");if(a)b+=c+f.format(e,a);var h=d.create(b,{name:"Sys.ArgumentNullException",paramName:a});h.popStackFrame();return h};d.argumentOutOfRange=function(i,g,j){var h="Sys.ArgumentOutOfRangeException: "+(j?j:"Specified argument was out of the range of valid values.");if(i)h+=c+f.format(e,i);if(typeof g!==b&&g!==a)h+=c+f.format("Actual value was {0}.",g);var k=d.create(h,{name:"Sys.ArgumentOutOfRangeException",paramName:i,actualValue:g});k.popStackFrame();return k};d.argumentType=function(h,g,b,i){var a="Sys.ArgumentTypeException: ";if(i)a+=i;else if(g&&b)a+=f.format("Object of type '{0}' cannot be converted to type '{1}'.",g.getName?g.getName():g,b.getName?b.getName():b);else a+="Object cannot be converted to the required type.";if(h)a+=c+f.format(e,h);var j=d.create(a,{name:"Sys.ArgumentTypeException",paramName:h,actualType:g,expectedType:b});j.popStackFrame();return j};d.argumentUndefined=function(a,g){var b="Sys.ArgumentUndefinedException: "+(g?g:"Value cannot be undefined.");if(a)b+=c+f.format(e,a);var h=d.create(b,{name:"Sys.ArgumentUndefinedException",paramName:a});h.popStackFrame();return h};d.invalidOperation=function(a){var c="Sys.InvalidOperationException: "+(a?a:"Operation is not valid due to the current state of the object."),b=d.create(c,{name:"Sys.InvalidOperationException"});b.popStackFrame();return b};return d}();g.MsAjaxError=h;var f=function(){function a(){}a.format=function(c){for(var b=[],a=1;a<arguments.length;a++)b[a-1]=arguments[a];var d=c;return d.replace(/{(\d+)}/gm,function(d,a){var c=parseInt(a,10);return b[c]===undefined?"{"+a+"}":b[c]})};a.startsWith=function(b,a){return b.substr(0,a.length)===a};return a}();g.MsAjaxString=f;var i=function(){function a(){}a.trace=function(){};return a}();g.MsAjaxDebug=i;if(!OsfMsAjaxFactory.isMsAjaxLoaded()){var e=function(a,c,b){if(a.__typeName===undefined)a.__typeName=c;if(a.__class===undefined)a.__class=b};e(Function,"Function",d);e(Error,"Error",d);e(Object,"Object",d);e(String,"String",d);e(Boolean,"Boolean",d);e(Date,"Date",d);e(Number,"Number",d);e(RegExp,"RegExp",d);e(Array,"Array",d);if(!Function.createCallback)Function.createCallback=function(b,a){var c=Function._validateParams(arguments,[{name:"method",type:Function},{name:"context",mayBeNull:d}]);if(c)throw c;return function(){var e=arguments.length;if(e>0){for(var d=[],c=0;c<e;c++)d[c]=arguments[c];d[e]=a;return b.apply(this,d)}return b.call(this,a)}};if(!Function.createDelegate)Function.createDelegate=function(b,c){var a=Function._validateParams(arguments,[{name:"instance",mayBeNull:d},{name:"method",type:Function}]);if(a)throw a;return function(){return c.apply(b,arguments)}};if(!Function._validateParams)Function._validateParams=function(i,g,e){var c,f=g.length;e=e||typeof e===b;c=Function._validateParameterCount(i,g,e);if(c){c.popStackFrame();return c}for(var d=0,k=i.length;d<k;d++){var h=g[Math.min(d,f-1)],j=h.name;if(h.parameterArray)j+="["+(d-f+1)+"]";else if(!e&&d>=f)break;c=Function._validateParameter(i[d],h,j);if(c){c.popStackFrame();return c}}return a};if(!Function._validateParameterCount)Function._validateParameterCount=function(m,f,l){var b,e,c=f.length,g=m.length;if(g<c){var i=c;for(b=0;b<c;b++){var j=f[b];if(j.optional||j.parameterArray)i--}if(g<i)e=d}else if(l&&g>c){e=d;for(b=0;b<c;b++)if(f[b].parameterArray){e=false;break}}if(e){var k=h.parameterCount();k.popStackFrame();return k}return a};if(!Function._validateParameter)Function._validateParameter=function(e,c,j){var d,i=c.type,n=!!c.integer,m=!!c.domElement,o=!!c.mayBeNull;d=Function._validateParameterType(e,i,n,m,o,j);if(d){d.popStackFrame();return d}var g=c.elementType,h=!!c.elementMayBeNull;if(i===Array&&typeof e!==b&&e!==a&&(g||!h))for(var l=!!c.elementInteger,k=!!c.elementDomElement,f=0;f<e.length;f++){var p=e[f];d=Function._validateParameterType(p,g,l,k,h,j+"["+f+"]");if(d){d.popStackFrame();return d}}return a};if(!Function._validateParameterType)Function._validateParameterType=function(d,e,j,i,h,f){var c,k;if(typeof d===b)if(h)return a;else{c=g.MsAjaxError.argumentUndefined(f);c.popStackFrame();return c}if(d===a)if(h)return a;else{c=g.MsAjaxError.argumentNull(f);c.popStackFrame();return c}if(e&&!g.MsAjaxTypeHelper.isInstanceOfType(e,d)){c=g.MsAjaxError.argumentType(f,typeof d,e);c.popStackFrame();return c}return a};if(!window.Type)window.Type=Function;if(!Type.registerNamespace)Type.registerNamespace=function(d){for(var c=d.split("."),b=window,a=0;a<c.length;a++){b[c[a]]=b[c[a]]||{};b=b[c[a]]}};if(!Type.prototype.registerClass)Type.prototype.registerClass=function(a){a={}};typeof Sys===b&&Type.registerNamespace("Sys");if(!Error.prototype.popStackFrame)Error.prototype.popStackFrame=function(){var d=this;if(arguments.length!==0)throw h.parameterCount();if(typeof d.stack===b||d.stack===a||typeof d.fileName===b||d.fileName===a||typeof d.lineNumber===b||d.lineNumber===a)return;var e=d.stack.split(c),g=e[0],j=d.fileName+":"+d.lineNumber;while(typeof g!==b&&g!==a&&g.indexOf(j)===-1){e.shift();g=e[0]}var i=e[1];if(typeof i===b||i===a)return;var f=i.match(/@(.*):(\d+)$/);if(typeof f===b||f===a)return;d.fileName=f[1];d.lineNumber=parseInt(f[2]);e.shift();d.stack=e.join(c)};OsfMsAjaxFactory.msAjaxError=h;OsfMsAjaxFactory.msAjaxString=f;OsfMsAjaxFactory.msAjaxDebug=i}})(OfficeExt||(OfficeExt={}));OSF.OUtil.setNamespace("SafeArray",OSF.DDA);OSF.DDA.SafeArray.Response={Status:0,Payload:1};OSF.DDA.SafeArray.UniqueArguments={Offset:"offset",Run:"run",BindingSpecificData:"bindingSpecificData",MergedCellGuid:"{66e7831f-81b2-42e2-823c-89e872d541b3}"};OSF.OUtil.setNamespace("Delegate",OSF.DDA.SafeArray);OSF.DDA.SafeArray.Delegate._onException=function(d,c){var a,b=d.number;if(b)switch(b){case -2146828218:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;break;case -2147467259:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;break;case -2146828283:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;break;case -2147209089:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;break;case -2146827850:default:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError}c.onComplete&&c.onComplete(a||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)};OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod=function(c){var a,b=c.number;if(b)switch(b){case -2146828218:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;break;case -2146827850:default:a=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError}return a||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError};OSF.DDA.SafeArray.Delegate.SpecialProcessor=function(){function a(a){var b;try{var h=a.ubound(1),d=a.ubound(2);a=a.toArray();if(h==1&&d==1)b=[a];else{b=[];for(var f=0;f<h;f++){for(var c=[],e=0;e<d;e++){var g=a[f*d+e];g!=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid&&c.push(g)}c.length>0&&b.push(c)}}}catch(i){}return b}var c=[],b={};b[Microsoft.Office.WebExtension.Parameters.Data]=function(){var c=0,b=1;return {toHost:function(a){if(OSF.DDA.TableDataProperties&&typeof a!="string"&&a[OSF.DDA.TableDataProperties.TableRows]!==undefined){var d=[];d[c]=a[OSF.DDA.TableDataProperties.TableRows];d[b]=a[OSF.DDA.TableDataProperties.TableHeaders];a=d}return a},fromHost:function(f){var e;if(f.toArray){var g=f.dimensions();if(g===2)e=a(f);else{var d=f.toArray();if(d.length===2&&(d[0]!=null&&d[0].toArray||d[1]!=null&&d[1].toArray)){e={};e[OSF.DDA.TableDataProperties.TableRows]=a(d[c]);e[OSF.DDA.TableDataProperties.TableHeaders]=a(d[b])}else e=d}}else e=f;return e}}}();OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this,c,b);this.unpack=function(c,a){var d;if(this.isComplexType(c)||OSF.DDA.ListType.isListType(c)){var e=(a||typeof a==="unknown")&&a.toArray;d=e?a.toArray():a||{}}else if(this.isDynamicType(c))d=b[c].fromHost(a);else d=a;return d}};OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor,OSF.DDA.SpecialProcessor);OSF.DDA.SafeArray.Delegate.ParameterMap=OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor,[{type:Microsoft.Office.WebExtension.Parameters.ValueFormat,toHost:[{name:Microsoft.Office.WebExtension.ValueFormat.Unformatted,value:0},{name:Microsoft.Office.WebExtension.ValueFormat.Formatted,value:1}]},{type:Microsoft.Office.WebExtension.Parameters.FilterType,toHost:[{name:Microsoft.Office.WebExtension.FilterType.All,value:0}]}]);OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.AsyncResultStatus,fromHost:[{name:Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded,value:0},{name:Microsoft.Office.WebExtension.AsyncResultStatus.Failed,value:1}]});OSF.DDA.SafeArray.Delegate.executeAsync=function(a){function c(a){var b=a;if(OSF.OUtil.isArray(a))for(var f=b.length,d=0;d<f;d++)b[d]=c(b[d]);else if(OSF.OUtil.isDate(a))b=a.getVarDate();else if(typeof a==="object"&&!OSF.OUtil.isArray(a)){b=[];for(var e in a)if(!OSF.OUtil.isFunction(a[e]))b[e]=c(a[e])}return b}function b(a){var e=a;if(a!=null&&a.toArray){var d=a.toArray();e=new Array(d.length);for(var c=0;c<d.length;c++)e[c]=b(d[c])}return e}try{a.onCalling&&a.onCalling();OSF.ClientHostController.execute(a.dispId,c(a.hostCallArgs),function(h){var d=h.toArray(),e=d[OSF.DDA.SafeArray.Response.Status];if(e==OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult){var c=d[OSF.DDA.SafeArray.Response.Payload];c=b(c);if(c!=null){if(!a._chunkResultData)a._chunkResultData=[];a._chunkResultData[c[0]]=c[1]}return false}a.onReceiving&&a.onReceiving();if(a.onComplete){var c;if(e==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess){if(d.length>2){c=[];for(var f=1;f<d.length;f++)c[f-1]=d[f]}else c=d[OSF.DDA.SafeArray.Response.Payload];if(a._chunkResultData){c=b(c);if(c!=null){var g=c[c.length-1];if(a._chunkResultData.length==g)c[c.length-1]=a._chunkResultData;else e=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError}}}else c=d[OSF.DDA.SafeArray.Response.Payload];a.onComplete(e,c)}return true})}catch(d){OSF.DDA.SafeArray.Delegate._onException(d,a)}};OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function(c,a){var b=(new Date).getTime();return function(d){a.onReceiving&&a.onReceiving();var e=d.toArray?d.toArray()[OSF.DDA.SafeArray.Response.Status]:d;a.onComplete&&a.onComplete(e);OSF.AppTelemetry&&OSF.AppTelemetry.onRegisterDone(c,a.dispId,Math.abs((new Date).getTime()-b),e)}};OSF.DDA.SafeArray.Delegate.registerEventAsync=function(a){a.onCalling&&a.onCalling();var c=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true,a);try{OSF.ClientHostController.registerEvent(a.dispId,a.targetId,function(c,b){a.onEvent&&a.onEvent(b);OSF.AppTelemetry&&OSF.AppTelemetry.onEventDone(a.dispId)},c)}catch(b){OSF.DDA.SafeArray.Delegate._onException(b,a)}};OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function(a){a.onCalling&&a.onCalling();var c=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false,a);try{OSF.ClientHostController.unregisterEvent(a.dispId,a.targetId,c)}catch(b){OSF.DDA.SafeArray.Delegate._onException(b,a)}};OSF.ClientMode={ReadWrite:0,ReadOnly:1};OSF.DDA.RichInitializationReason={1:Microsoft.Office.WebExtension.InitializationReason.Inserted,2:Microsoft.Office.WebExtension.InitializationReason.DocumentOpened};OSF.InitializationHelper=function(d,b,f,e,c){var a=this;a._hostInfo=d;a._webAppState=b;a._context=f;a._settings=e;a._hostFacade=c;a._initializeSettings=a.initializeSettings};OSF.InitializationHelper.prototype.deserializeSettings=function(b,f){var d,c=OSF.OUtil.getSessionStorage();if(c){var a=c.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());if(a)b=JSON.parse(a);else{a=JSON.stringify(b);c.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(),a)}}var e=OSF.DDA.SettingsManager.deserializeSettings(b);if(f)d=new OSF.DDA.RefreshableSettings(e);else d=new OSF.DDA.Settings(e);return d};OSF.InitializationHelper.prototype.saveAndSetDialogInfo=function(){};OSF.InitializationHelper.prototype.setAgaveHostCommunication=function(){};OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize=function(a){this.prepareApiSurface(a);Microsoft.Office.WebExtension.initialize(this.getInitializationReason(a))};OSF.InitializationHelper.prototype.prepareApiSurface=function(a){var e=new OSF.DDA.License(a.get_eToken()),d=OSF.DDA.OfficeTheme&&OSF.DDA.OfficeTheme.getOfficeTheme?OSF.DDA.OfficeTheme.getOfficeTheme:null;if(a.get_isDialog()){if(OSF.DDA.UI.ChildUI)a.ui=new OSF.DDA.UI.ChildUI}else if(OSF.DDA.UI.ParentUI){a.ui=new OSF.DDA.UI.ParentUI;OfficeExt.Container&&OSF.DDA.DispIdHost.addAsyncMethods(a.ui,[OSF.DDA.AsyncMethodNames.CloseContainerAsync])}OSF.DDA.OpenBrowser&&OSF.DDA.DispIdHost.addAsyncMethods(a.ui,[OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);if(OSF.DDA.Auth){a.auth=new OSF.DDA.Auth;OSF.DDA.DispIdHost.addAsyncMethods(a.auth,[OSF.DDA.AsyncMethodNames.GetAccessTokenAsync])}OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(a,a.doc,e,null,d));var b,c;b=OSF.DDA.DispIdHost.getClientDelegateMethods;c=OSF.DDA.SafeArray.Delegate.ParameterMap;OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(b,c))};OSF.InitializationHelper.prototype.getInitializationReason=function(a){return OSF.DDA.RichInitializationReason[a.get_reason()]};OSF.DDA.DispIdHost.getClientDelegateMethods=function(b){var a={};a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;a[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;a[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;a[OSF.DDA.DispIdHost.Delegates.OpenDialog]=OSF.DDA.SafeArray.Delegate.openDialog;a[OSF.DDA.DispIdHost.Delegates.CloseDialog]=OSF.DDA.SafeArray.Delegate.closeDialog;a[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.messageParent;a[OSF.DDA.DispIdHost.Delegates.SendMessage]=OSF.DDA.SafeArray.Delegate.sendMessage;if(OSF.DDA.AsyncMethodNames.RefreshAsync&&b==OSF.DDA.AsyncMethodNames.RefreshAsync.id){var d=function(c,b,a){return OSF.DDA.ClientSettingsManager.read(b,a)};a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(d)}if(OSF.DDA.AsyncMethodNames.SaveAsync&&b==OSF.DDA.AsyncMethodNames.SaveAsync.id){var c=function(a,c,b){return OSF.DDA.ClientSettingsManager.write(a[OSF.DDA.SettingsManager.SerializedSettings],a[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale],c,b)};a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(c)}return a};var OfficeExt;(function(b){var a=function(){function a(){}a.prototype.execute=function(c,b,a){window.external.Execute(c,b,a)};a.prototype.registerEvent=function(d,b,c,a){window.external.RegisterEvent(d,b,c,a)};a.prototype.unregisterEvent=function(c,b,a){window.external.UnregisterEvent(c,b,a)};return a}();b.RichClientHostController=a})(OfficeExt||(OfficeExt={}));var OfficeExt;(function(a){var b=function(b){__extends(a,b);function a(){b.apply(this,arguments)}a.prototype.messageParent=function(b){var a=b[Microsoft.Office.WebExtension.Parameters.MessageToParent];window.external.MessageParent(a)};a.prototype.openDialog=function(d,b,c,a){this.registerEvent(d,b,c,a)};a.prototype.closeDialog=function(c,b,a){this.unregisterEvent(c,b,a)};a.prototype.sendMessage=function(){};return a}(a.RichClientHostController);a.Win32RichClientHostController=b})(OfficeExt||(OfficeExt={}));OSF.ClientHostController=new OfficeExt.Win32RichClientHostController;var OfficeExt;(function(a){var b;(function(c){var b=function(){var a=null;function b(){this._osfOfficeTheme=a;this._osfOfficeThemeTimeStamp=a}b.prototype.getOfficeTheme=function(){var c="GetOfficeThemeInfo",a=this;if(OSF.DDA._OsfControlContext){if(a._osfOfficeTheme&&a._osfOfficeThemeTimeStamp&&(new Date).getTime()-a._osfOfficeThemeTimeStamp<b._osfOfficeThemeCacheValidPeriod)OSF.AppTelemetry&&OSF.AppTelemetry.onPropertyDone(c,0);else{var g=(new Date).getTime(),f=OSF.DDA._OsfControlContext.GetOfficeThemeInfo(),d=(new Date).getTime();OSF.AppTelemetry&&OSF.AppTelemetry.onPropertyDone(c,Math.abs(d-g));a._osfOfficeTheme=JSON.parse(f);for(var e in a._osfOfficeTheme)a._osfOfficeTheme[e]=OSF.OUtil.convertIntToCssHexColor(a._osfOfficeTheme[e]);a._osfOfficeThemeTimeStamp=d}return a._osfOfficeTheme}};b.instance=function(){if(b._instance==a)b._instance=new b;return b._instance};b._osfOfficeThemeCacheValidPeriod=5e3;b._instance=a;return b}();c.OfficeThemeManager=b;OSF.OUtil.setNamespace("OfficeTheme",OSF.DDA);OSF.DDA.OfficeTheme.getOfficeTheme=a.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme})(b=a.OfficeTheme||(a.OfficeTheme={}))})(OfficeExt||(OfficeExt={}));OSF.DDA.ClientSettingsManager={getSettingsExecuteMethod:function(a){return function(b){var d,c;try{c=a(b.hostCallArgs,b.onCalling,b.onReceiving);d=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess}catch(e){d=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;c={name:Strings.OfficeOM.L_InternalError,message:e}}b.onComplete&&b.onComplete(d,c)}},read:function(e,d){var b=[],f=[];e&&e();OSF.DDA._OsfControlContext.GetSettings().Read(b,f);d&&d();for(var c={},a=0;a<b.length;a++)c[b[a]]=f[a];return c},write:function(a,g,c,b){var e=[],d=[];for(var f in a){e.push(f);d.push(a[f])}c&&c();OSF.DDA._OsfControlContext.GetSettings().Write(e,d);b&&b()}};OSF.InitializationHelper.prototype.initializeSettings=function(b){var a=OSF.DDA.ClientSettingsManager.read(),c=this.deserializeSettings(a,b);return c};OSF.InitializationHelper.prototype.getAppContext=function(B,o){var b="undefined",d,a,m="Warning: Office.js is loaded outside of Office client";try{if(window.external&&typeof window.external.GetContext!==b)a=OSF.DDA._OsfControlContext=window.external.GetContext();else{OsfMsAjaxFactory.msAjaxDebug.trace(m);return}}catch(A){OsfMsAjaxFactory.msAjaxDebug.trace(m);return}var u=a.GetAppType(),z=a.GetSolutionRef(),v=a.GetAppVersionMajor(),p=a.GetAppVersionMinor(),t=a.GetAppUILocale(),r=a.GetAppDataLocale(),w=a.GetDocUrl(),q=a.GetAppCapabilities(),x=a.GetActivationMode(),n=a.GetControlIntegrationLevel(),s=[],c;try{c=a.GetSolutionToken()}catch(y){}var k;if(typeof a.GetCorrelationId!==b)k=a.GetCorrelationId();var j;if(typeof a.GetInstanceId!==b)j=a.GetInstanceId();var l;if(typeof a.GetTouchEnabled!==b)l=a.GetTouchEnabled();var h;if(typeof a.GetCommerceAllowed!==b)h=a.GetCommerceAllowed();var g;if(typeof a.GetSupportedMatrix!==b)g=a.GetSupportedMatrix();var f;if(typeof a.GetHostCustomMessage!==b)f=a.GetHostCustomMessage();var i;if(typeof a.GetHostFullVersion!==b)i=a.GetHostFullVersion();var e;if(typeof a.GetDialogRequirementMatrix!=b)e=a.GetDialogRequirementMatrix();c=c?c.toString():"";d=new OSF.OfficeAppContext(z,u,v,t,r,w,q,s,x,n,c,k,j,l,h,p,g,f,i,undefined,undefined,undefined,e);OSF.AppTelemetry&&OSF.AppTelemetry.initialize(d);o(d)};var OSFLog;(function(g){var e="ResponseTime",d="Message",c="SessionId",b="CorrelationId",a=true,f=function(){function b(a){this._table=a;this._fields={}}Object.defineProperty(b.prototype,"Fields",{"get":function(){return this._fields},enumerable:a,configurable:a});Object.defineProperty(b.prototype,"Table",{"get":function(){return this._table},enumerable:a,configurable:a});b.prototype.SerializeFields=function(){};b.prototype.SetSerializedField=function(b,a){if(typeof a!=="undefined"&&a!==null)this._serializedFields[b]=a.toString()};b.prototype.SerializeRow=function(){var a=this;a._serializedFields={};a.SetSerializedField("Table",a._table);a.SerializeFields();return JSON.stringify(a._serializedFields)};return b}();g.BaseUsageData=f;var i=function(v){var u="IsFromWacAutomation",t="WacHostEnvironment",s="HostJSVersion",r="OfficeJSVersion",q="DocUrl",p="AppSizeHeight",o="AppSizeWidth",n="ClientId",m="HostVersion",l="Host",k="UserId",j="Browser",i="AssetId",h="AppURL",g="AppInstanceId",f="AppId";__extends(e,v);function e(){v.call(this,"AppActivated")}Object.defineProperty(e.prototype,b,{"get":function(){return this.Fields[b]},"set":function(a){this.Fields[b]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,c,{"get":function(){return this.Fields[c]},"set":function(a){this.Fields[c]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,f,{"get":function(){return this.Fields[f]},"set":function(a){this.Fields[f]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,g,{"get":function(){return this.Fields[g]},"set":function(a){this.Fields[g]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,h,{"get":function(){return this.Fields[h]},"set":function(a){this.Fields[h]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,i,{"get":function(){return this.Fields[i]},"set":function(a){this.Fields[i]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,j,{"get":function(){return this.Fields[j]},"set":function(a){this.Fields[j]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,k,{"get":function(){return this.Fields[k]},"set":function(a){this.Fields[k]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,l,{"get":function(){return this.Fields[l]},"set":function(a){this.Fields[l]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,m,{"get":function(){return this.Fields[m]},"set":function(a){this.Fields[m]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,n,{"get":function(){return this.Fields[n]},"set":function(a){this.Fields[n]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,o,{"get":function(){return this.Fields[o]},"set":function(a){this.Fields[o]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,p,{"get":function(){return this.Fields[p]},"set":function(a){this.Fields[p]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,d,{"get":function(){return this.Fields[d]},"set":function(a){this.Fields[d]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,q,{"get":function(){return this.Fields[q]},"set":function(a){this.Fields[q]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,r,{"get":function(){return this.Fields[r]},"set":function(a){this.Fields[r]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,s,{"get":function(){return this.Fields[s]},"set":function(a){this.Fields[s]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,t,{"get":function(){return this.Fields[t]},"set":function(a){this.Fields[t]=a},enumerable:a,configurable:a});Object.defineProperty(e.prototype,u,{"get":function(){return this.Fields[u]},"set":function(a){this.Fields[u]=a},enumerable:a,configurable:a});e.prototype.SerializeFields=function(){var a=this;a.SetSerializedField(b,a.CorrelationId);a.SetSerializedField(c,a.SessionId);a.SetSerializedField(f,a.AppId);a.SetSerializedField(g,a.AppInstanceId);a.SetSerializedField(h,a.AppURL);a.SetSerializedField(i,a.AssetId);a.SetSerializedField(j,a.Browser);a.SetSerializedField(k,a.UserId);a.SetSerializedField(l,a.Host);a.SetSerializedField(m,a.HostVersion);a.SetSerializedField(n,a.ClientId);a.SetSerializedField(o,a.AppSizeWidth);a.SetSerializedField(p,a.AppSizeHeight);a.SetSerializedField(d,a.Message);a.SetSerializedField(q,a.DocUrl);a.SetSerializedField(r,a.OfficeJSVersion);a.SetSerializedField(s,a.HostJSVersion);a.SetSerializedField(t,a.WacHostEnvironment);a.SetSerializedField(u,a.IsFromWacAutomation)};return e}(f);g.AppActivatedUsageData=i;var j=function(h){var f="StartTime",d="ScriptId";__extends(g,h);function g(){h.call(this,"ScriptLoad")}Object.defineProperty(g.prototype,b,{"get":function(){return this.Fields[b]},"set":function(a){this.Fields[b]=a},enumerable:a,configurable:a});Object.defineProperty(g.prototype,c,{"get":function(){return this.Fields[c]},"set":function(a){this.Fields[c]=a},enumerable:a,configurable:a});Object.defineProperty(g.prototype,d,{"get":function(){return this.Fields[d]},"set":function(a){this.Fields[d]=a},enumerable:a,configurable:a});Object.defineProperty(g.prototype,f,{"get":function(){return this.Fields[f]},"set":function(a){this.Fields[f]=a},enumerable:a,configurable:a});Object.defineProperty(g.prototype,e,{"get":function(){return this.Fields[e]},"set":function(a){this.Fields[e]=a},enumerable:a,configurable:a});g.prototype.SerializeFields=function(){var a=this;a.SetSerializedField(b,a.CorrelationId);a.SetSerializedField(c,a.SessionId);a.SetSerializedField(d,a.ScriptId);a.SetSerializedField(f,a.StartTime);a.SetSerializedField(e,a.ResponseTime)};return g}(f);g.ScriptLoadUsageData=j;var k=function(j){var h="CloseMethod",g="OpenTime",f="AppSizeFinalHeight",e="AppSizeFinalWidth",d="FocusTime";__extends(i,j);function i(){j.call(this,"AppClosed")}Object.defineProperty(i.prototype,b,{"get":function(){return this.Fields[b]},"set":function(a){this.Fields[b]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,c,{"get":function(){return this.Fields[c]},"set":function(a){this.Fields[c]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,d,{"get":function(){return this.Fields[d]},"set":function(a){this.Fields[d]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,e,{"get":function(){return this.Fields[e]},"set":function(a){this.Fields[e]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,f,{"get":function(){return this.Fields[f]},"set":function(a){this.Fields[f]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,g,{"get":function(){return this.Fields[g]},"set":function(a){this.Fields[g]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,h,{"get":function(){return this.Fields[h]},"set":function(a){this.Fields[h]=a},enumerable:a,configurable:a});i.prototype.SerializeFields=function(){var a=this;a.SetSerializedField(b,a.CorrelationId);a.SetSerializedField(c,a.SessionId);a.SetSerializedField(d,a.FocusTime);a.SetSerializedField(e,a.AppSizeFinalWidth);a.SetSerializedField(f,a.AppSizeFinalHeight);a.SetSerializedField(g,a.OpenTime);a.SetSerializedField(h,a.CloseMethod)};return i}(f);g.AppClosedUsageData=k;var l=function(j){var h="ErrorType",g="Parameters",f="APIID",d="APIType";__extends(i,j);function i(){j.call(this,"APIUsage")}Object.defineProperty(i.prototype,b,{"get":function(){return this.Fields[b]},"set":function(a){this.Fields[b]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,c,{"get":function(){return this.Fields[c]},"set":function(a){this.Fields[c]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,d,{"get":function(){return this.Fields[d]},"set":function(a){this.Fields[d]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,f,{"get":function(){return this.Fields[f]},"set":function(a){this.Fields[f]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,g,{"get":function(){return this.Fields[g]},"set":function(a){this.Fields[g]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,e,{"get":function(){return this.Fields[e]},"set":function(a){this.Fields[e]=a},enumerable:a,configurable:a});Object.defineProperty(i.prototype,h,{"get":function(){return this.Fields[h]},"set":function(a){this.Fields[h]=a},enumerable:a,configurable:a});i.prototype.SerializeFields=function(){var a=this;a.SetSerializedField(b,a.CorrelationId);a.SetSerializedField(c,a.SessionId);a.SetSerializedField(d,a.APIType);a.SetSerializedField(f,a.APIID);a.SetSerializedField(g,a.Parameters);a.SetSerializedField(e,a.ResponseTime);a.SetSerializedField(h,a.ErrorType)};return i}(f);g.APIUsageUsageData=l;var h=function(g){var e="SuccessCode";__extends(f,g);function f(){g.call(this,"AppInitialization")}Object.defineProperty(f.prototype,b,{"get":function(){return this.Fields[b]},"set":function(a){this.Fields[b]=a},enumerable:a,configurable:a});Object.defineProperty(f.prototype,c,{"get":function(){return this.Fields[c]},"set":function(a){this.Fields[c]=a},enumerable:a,configurable:a});Object.defineProperty(f.prototype,e,{"get":function(){return this.Fields[e]},"set":function(a){this.Fields[e]=a},enumerable:a,configurable:a});Object.defineProperty(f.prototype,d,{"get":function(){return this.Fields[d]},"set":function(a){this.Fields[d]=a},enumerable:a,configurable:a});f.prototype.SerializeFields=function(){var a=this;a.SetSerializedField(b,a.CorrelationId);a.SetSerializedField(c,a.SessionId);a.SetSerializedField(e,a.SuccessCode);a.SetSerializedField(d,a.Message)};return f}(f);g.AppInitializationUsageData=h})(OSFLog||(OSFLog={}));var Logger;(function(a){"use strict";(function(a){a[a["info"]=0]="info";a[a["warning"]=1]="warning";a[a["error"]=2]="error"})(a.TraceLevel||(a.TraceLevel={}));var f=a.TraceLevel;(function(a){a[a["none"]=0]="none";a[a["flush"]=1]="flush"})(a.SendFlag||(a.SendFlag={}));var g=a.SendFlag;function b(){OSF.Logger&&OSF.Logger.ulsEndpoint&&OSF.Logger.ulsEndpoint.loadProxyFrame()}a.allowUploadingData=b;function e(a,c,d){if(OSF.Logger&&OSF.Logger.ulsEndpoint){var b={traceLevel:a,message:c,flag:d,internalLog:true},e=JSON.stringify(b);OSF.Logger.ulsEndpoint.writeLog(e)}}a.sendLog=e;function c(){try{return new d}catch(a){return null}}var d=function(){function a(){var a=this,b=a;a.proxyFrame=null;a.telemetryEndPoint="https://telemetryservice.firstpartyapps.oaspapps.com/telemetryservice/telemetryproxy.html";a.buffer=[];a.proxyFrameReady=false;OSF.OUtil.addEventListener(window,"message",function(a){return b.tellProxyFrameReady(a)});setTimeout(function(){b.loadProxyFrame()},3e3)}a.prototype.writeLog=function(c){var b=this;if(b.proxyFrameReady===true)b.proxyFrame.contentWindow.postMessage(c,a.telemetryOrigin);else b.buffer.length<128&&b.buffer.push(c)};a.prototype.loadProxyFrame=function(){var a=this;if(a.proxyFrame==null){a.proxyFrame=document.createElement("iframe");a.proxyFrame.setAttribute("style","display:none");a.proxyFrame.setAttribute("src",a.telemetryEndPoint);document.head.appendChild(a.proxyFrame)}};a.prototype.tellProxyFrameReady=function(d){var b=this,g=b;if(d.data==="ProxyFrameReadyToLog"){b.proxyFrameReady=true;for(var c=0;c<b.buffer.length;c++)b.writeLog(b.buffer[c]);b.buffer.length=0;OSF.OUtil.removeEventListener(window,"message",function(a){return g.tellProxyFrameReady(a)})}else if(d.data==="ProxyFrameReadyToInit"){var e={appName:"Office APPs",sessionId:OSF.OUtil.Guid.generateNewGuid()},f=JSON.stringify(e);b.proxyFrame.contentWindow.postMessage(f,a.telemetryOrigin)}};a.telemetryOrigin="https://telemetryservice.firstpartyapps.oaspapps.com";return a}();if(!OSF.Logger)OSF.Logger=a;a.ulsEndpoint=c()})(Logger||(Logger={}));var OSFAriaLogger;(function(a){var b=function(){function a(){}a.prototype.getAriaCDNLocation=function(){return OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath()+"/ariatelemetry/aria-web-telemetry.js"};a.getInstance=function(){if(a.AriaLoggerObj===undefined)a.AriaLoggerObj=new a;return a.AriaLoggerObj};a.prototype.isIUsageData=function(a){return a["Fields"]!==undefined};a.prototype.loadAriaScriptAndLog=function(c,a){var b=1e3;OSF.OUtil.loadScript(this.getAriaCDNLocation(),function(){try{if(!this.ALogger){var e="db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";this.ALogger=AWTLogManager.initialize(e)}var b=new AWTEventProperties;b.setName("Office.Extensibility.OfficeJS."+c);for(var d in a)d.toLowerCase()!=="table"&&b.setProperty(d,a[d]);var f=new Date;b.setProperty("Date",f.toISOString());this.ALogger.logEvent(b)}catch(g){}},b)};a.prototype.logData=function(a){if(this.isIUsageData(a))this.loadAriaScriptAndLog(a["Table"],a["Fields"]);else this.loadAriaScriptAndLog(a["Table"],a)};return a}();a.AriaLogger=b})(OSFAriaLogger||(OSFAriaLogger={}));var OSFAppTelemetry;(function(d){var b=null,e=true,c="";"use strict";var a,g=OSF.OUtil.Guid.generateNewGuid(),j=c,p=new RegExp("^https?://store\\.office(ppe|-int)?\\.com/","i");d.enableTelemetry=e;var z=function(){function a(){}return a}(),h=function(){function a(b,a){this.name=b;this.handler=a}return a}(),l=function(){function a(){this.clientIDKey="Office API client";this.logIdSetKey="Office App Log Id Set"}a.prototype.getClientId=function(){var b=this,a=b.getValue(b.clientIDKey);if(!a||a.length<=0||a.length>40){a=OSF.OUtil.Guid.generateNewGuid();b.setValue(b.clientIDKey,a)}return a};a.prototype.saveLog=function(d,e){var b=this,a=b.getValue(b.logIdSetKey);a=(a&&a.length>0?a+";":c)+d;b.setValue(b.logIdSetKey,a);b.setValue(d,e)};a.prototype.enumerateLog=function(c,e){var a=this,d=a.getValue(a.logIdSetKey);if(d){var f=d.split(";");for(var h in f){var b=f[h],g=a.getValue(b);if(g){c&&c(b,g);e&&a.remove(b)}}e&&a.remove(a.logIdSetKey)}};a.prototype.getValue=function(d){var a=OSF.OUtil.getLocalStorage(),b=c;if(a)b=a.getItem(d);return b};a.prototype.setValue=function(c,b){var a=OSF.OUtil.getLocalStorage();a&&a.setItem(c,b)};a.prototype.remove=function(b){var a=OSF.OUtil.getLocalStorage();if(a)try{a.removeItem(b)}catch(c){}};return a}(),i=function(){function a(){}a.prototype.LogData=function(a){if(!OSF.Logger||!d.enableTelemetry)return;try{OSFAriaLogger.AriaLogger.getInstance().logData(a)}catch(b){}};a.prototype.LogRawData=function(a){if(!OSF.Logger||!d.enableTelemetry)return;try{OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(a))}catch(b){}};return a}();function f(a){if(a)a=a.replace(/[{}]/g,c).toLowerCase();return a||c}function x(g){if(!OSF.Logger)return;if(a)return;a=new z;if(g.get_hostFullVersion())a.hostVersion=g.get_hostFullVersion();else a.hostVersion=g.get_appVersion();a.appId=g.get_id();a.host=g.get_appName();a.browser=window.navigator.userAgent;a.correlationId=f(g.get_correlationId());a.clientId=(new l).getClientId();a.appInstanceId=g.get_appInstanceId();if(a.appInstanceId)a.appInstanceId=a.appInstanceId.replace(/[{}]/g,c).toLowerCase();a.message=g.get_hostCustomMessage();a.officeJSVersion=OSF.ConstantNames.FileVersion;a.hostJSVersion="16.0.9009.1000";if(g._wacHostEnvironment)a.wacHostEnvironment=g._wacHostEnvironment;if(g._isFromWacAutomation!==undefined&&g._isFromWacAutomation!==b)a.isFromWacAutomation=g._isFromWacAutomation.toString().toLowerCase();var j=g.get_docUrl();a.docUrl=p.test(j)?j:c;var i=location.href;if(i)i=i.split("?")[0].split("#")[0];a.appURL=i;(function(i,a){var e,h,d;a.assetId=c;a.userId=c;try{e=decodeURIComponent(i);h=new DOMParser;d=h.parseFromString(e,"text/xml");var f=d.getElementsByTagName("t")[0].attributes.getNamedItem("cid"),g=d.getElementsByTagName("t")[0].attributes.getNamedItem("oid");if(f&&f.nodeValue)a.userId=f.nodeValue;else if(g&&g.nodeValue)a.userId=g.nodeValue;a.assetId=d.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue}catch(j){}finally{e=b;d=b;h=b}})(g.get_eToken(),a);(function(){var l=new Date,c=b,i=0,k=false,f=function(){if(document.hasFocus()){if(c==b)c=new Date}else if(c){i+=Math.abs((new Date).getTime()-c.getTime());c=b}},a=[];a.push(new h("focus",f));a.push(new h("blur",f));a.push(new h("focusout",f));a.push(new h("focusin",f));var j=function(){for(var f=0;f<a.length;f++)OSF.OUtil.removeEventListener(window,a[f].name,a[f].handler);a.length=0;if(!k){if(document.hasFocus()&&c){i+=Math.abs((new Date).getTime()-c.getTime());c=b}d.onAppClosed(Math.abs((new Date).getTime()-l.getTime()),i);k=e}};a.push(new h("beforeunload",j));a.push(new h("unload",j));for(var g=0;g<a.length;g++)OSF.OUtil.addEventListener(window,a[g].name,a[g].handler);f()})();d.onAppActivated()}d.initialize=x;function q(){if(!a)return;(new l).enumerateLog(function(b,a){return (new i).LogRawData(a)},e);var c=new OSFLog.AppActivatedUsageData;c.SessionId=g;c.AppId=a.appId;c.AssetId=a.assetId;c.AppURL=a.appURL;c.UserId=a.userId;c.ClientId=a.clientId;c.Browser=a.browser;c.Host=a.host;c.HostVersion=a.hostVersion;c.CorrelationId=f(a.correlationId);c.AppSizeWidth=window.innerWidth;c.AppSizeHeight=window.innerHeight;c.AppInstanceId=a.appInstanceId;c.Message=a.message;c.DocUrl=a.docUrl;c.OfficeJSVersion=a.officeJSVersion;c.HostJSVersion=a.hostJSVersion;if(a.wacHostEnvironment)c.WacHostEnvironment=a.wacHostEnvironment;if(a.isFromWacAutomation!==undefined&&a.isFromWacAutomation!==b)c.IsFromWacAutomation=a.isFromWacAutomation;(new i).LogData(c);setTimeout(function(){if(!OSF.Logger)return;OSF.Logger.allowUploadingData()},100)}d.onAppActivated=q;function u(e,d,c,b){var a=new OSFLog.ScriptLoadUsageData;a.CorrelationId=f(b);a.SessionId=g;a.ScriptId=e;a.StartTime=d;a.ResponseTime=c;(new i).LogData(a)}d.onScriptDone=u;function y(h,k,d,c,e){if(!a)return;var b=new OSFLog.APIUsageUsageData;b.CorrelationId=f(j);b.SessionId=g;b.APIType=h;b.APIID=k;b.Parameters=d;b.ResponseTime=c;b.ErrorType=e;(new i).LogData(b)}d.onCallDone=y;function t(h,d,f,g){var a=b;if(d)if(typeof d=="number")a=String(d);else if(typeof d==="object")for(var e in d){if(a!==b)a+=",";else a=c;if(typeof d[e]=="number")a+=String(d[e])}else a=c;OSF.AppTelemetry.onCallDone("method",h,a,f,g)}d.onMethodDone=t;function r(b,a){OSF.AppTelemetry.onCallDone("property",-1,b,a)}d.onPropertyDone=r;function w(c,a){OSF.AppTelemetry.onCallDone("event",c,b,0,a)}d.onEventDone=w;function s(d,e,a,c){OSF.AppTelemetry.onCallDone(d?"registerevent":"unregisterevent",e,b,a,c)}d.onRegisterDone=s;function v(d,c){if(!a)return;var b=new OSFLog.AppClosedUsageData;b.CorrelationId=f(j);b.SessionId=g;b.FocusTime=c;b.OpenTime=d;b.AppSizeFinalWidth=window.innerWidth;b.AppSizeFinalHeight=window.innerHeight;(new l).saveLog(g,b.SerializeRow())}d.onAppClosed=v;function m(a){j=f(a)}d.setOsfControlAppCorrelationId=m;function k(b,c){var a=new OSFLog.AppInitializationUsageData;a.CorrelationId=f(j);a.SessionId=g;a.SuccessCode=b?1:0;a.Message=c;(new i).LogData(a)}d.doAppInitializationLogging=k;function n(a){k(false,a)}d.logAppCommonMessage=n;function o(a){k(e,a)}d.logAppException=o;OSF.AppTelemetry=d})(OSFAppTelemetry||(OSFAppTelemetry={}));Microsoft.Office.WebExtension.BindingType={Table:"table",Text:"text",Matrix:"matrix"};OSF.DDA.BindingProperties={Id:"BindingId",Type:Microsoft.Office.WebExtension.Parameters.BindingType};OSF.OUtil.augmentList(OSF.DDA.ListDescriptors,{BindingList:"BindingList"});OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{Subset:"subset",BindingProperties:"BindingProperties"});OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList,OSF.DDA.PropertyDescriptors.BindingProperties);OSF.DDA.BindingPromise=function(b,a){this._id=b;OSF.OUtil.defineEnumerableProperty(this,"onFail",{"get":function(){return a},"set":function(c){var b=typeof c;if(b!="undefined"&&b!="function")throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction,b);a=c}})};OSF.DDA.BindingPromise.prototype={_fetch:function(b){var a=this;if(a.binding)b&&b(a.binding);else if(!a._binding){var c=a;Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(a._id,function(a){if(a.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded){OSF.OUtil.defineEnumerableProperty(c,"binding",{value:a.value});b&&b(c.binding)}else c.onFail&&c.onFail(a)})}return a},getDataAsync:function(){var a=arguments;this._fetch(function(b){b.getDataAsync.apply(b,a)});return this},setDataAsync:function(){var a=arguments;this._fetch(function(b){b.setDataAsync.apply(b,a)});return this},addHandlerAsync:function(){var a=arguments;this._fetch(function(b){b.addHandlerAsync.apply(b,a)});return this},removeHandlerAsync:function(){var a=arguments;this._fetch(function(b){b.removeHandlerAsync.apply(b,a)});return this}};OSF.DDA.BindingFacade=function(b){this._eventDispatches=[];OSF.OUtil.defineEnumerableProperty(this,"document",{value:b});var a=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[a.AddFromSelectionAsync,a.AddFromNamedItemAsync,a.GetAllAsync,a.GetByIdAsync,a.ReleaseByIdAsync])};OSF.DDA.UnknownBinding=function(b,a){OSF.OUtil.defineEnumerableProperties(this,{document:{value:a},id:{value:b}})};OSF.DDA.Binding=function(a,c){OSF.OUtil.defineEnumerableProperties(this,{document:{value:c},id:{value:a}});var d=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[d.GetDataAsync,d.SetDataAsync]);var e=Microsoft.Office.WebExtension.EventType,b=c.bindings._eventDispatches;if(!b[a])b[a]=new OSF.EventDispatch([e.BindingSelectionChanged,e.BindingDataChanged]);var f=b[a];OSF.DDA.DispIdHost.addEventSupport(this,f)};OSF.DDA.generateBindingId=function(){return "UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+(new Date).getTime()};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureBinding=function(a,c){var d=a[OSF.DDA.BindingProperties.Id],g=a[OSF.DDA.BindingProperties.RowCount],f=a[OSF.DDA.BindingProperties.ColumnCount],h=a[OSF.DDA.BindingProperties.HasHeaders],b;switch(a[OSF.DDA.BindingProperties.Type]){case Microsoft.Office.WebExtension.BindingType.Text:b=new OSF.DDA.TextBinding(d,c);break;case Microsoft.Office.WebExtension.BindingType.Matrix:b=new OSF.DDA.MatrixBinding(d,c,g,f);break;case Microsoft.Office.WebExtension.BindingType.Table:var i=function(){return OSF.DDA.ExcelDocument&&Microsoft.Office.WebExtension.context.document&&Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument},e;if(i()&&OSF.DDA.ExcelTableBinding)e=OSF.DDA.ExcelTableBinding;else e=OSF.DDA.TableBinding;b=new e(d,c,g,f,h);break;default:b=new OSF.DDA.UnknownBinding(d,c)}return b};OSF.DDA.AsyncMethodNames.addNames({AddFromSelectionAsync:"addFromSelectionAsync",AddFromNamedItemAsync:"addFromNamedItemAsync",GetAllAsync:"getAllAsync",GetByIdAsync:"getByIdAsync",ReleaseByIdAsync:"releaseByIdAsync",GetDataAsync:"getDataAsync",SetDataAsync:"setDataAsync"});(function(){var d="number",c="object",b="string",a=null;function e(a){return OSF.DDA.OMFactory.manufactureBinding(a,Microsoft.Office.WebExtension.context.document)}function f(a){return a.id}function g(c,e,d){var b=c[Microsoft.Office.WebExtension.Parameters.Data];if(OSF.DDA.TableDataProperties&&b&&(b[OSF.DDA.TableDataProperties.TableRows]!=undefined||b[OSF.DDA.TableDataProperties.TableHeaders]!=undefined))b=OSF.DDA.OMFactory.manufactureTableData(b);b=OSF.DDA.DataCoercion.coerceData(b,d[Microsoft.Office.WebExtension.Parameters.CoercionType]);return b==undefined?a:b}OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.BindingType,"enum":Microsoft.Office.WebExtension.BindingType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:{types:[b],calculate:OSF.DDA.generateBindingId}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[c],defaultValue:a}}],privateStateCallbacks:[],onSucceeded:e});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.ItemName,types:[b]},{name:Microsoft.Office.WebExtension.Parameters.BindingType,"enum":Microsoft.Office.WebExtension.BindingType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:{types:[b],calculate:OSF.DDA.generateBindingId}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[c],defaultValue:a}}],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.FailOnCollision,value:function(){return true}}],onSucceeded:e});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetAllAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(a){return OSF.OUtil.mapList(a[OSF.DDA.ListDescriptors.BindingList],e)}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[b]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:e});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[b]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(d,b,a){var c=a[Microsoft.Office.WebExtension.Parameters.Id];delete b._eventDispatches[c]}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataAsync,requiredArguments:[],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(b,a){return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(a.type)}}},{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:{"enum":Microsoft.Office.WebExtension.ValueFormat,defaultValue:Microsoft.Office.WebExtension.ValueFormat.Unformatted}},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:{"enum":Microsoft.Office.WebExtension.FilterType,defaultValue:Microsoft.Office.WebExtension.FilterType.All}},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:{types:[c,b],defaultValue:a}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[c],defaultValue:a}},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:{types:[d],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:{types:[d],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.RowCount,value:{types:[d],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.ColumnCount,value:{types:[d],defaultValue:0}}],checkCallArgs:function(a,b){if(a[Microsoft.Office.WebExtension.Parameters.StartRow]==0&&a[Microsoft.Office.WebExtension.Parameters.StartColumn]==0&&a[Microsoft.Office.WebExtension.Parameters.RowCount]==0&&a[Microsoft.Office.WebExtension.Parameters.ColumnCount]==0){delete a[Microsoft.Office.WebExtension.Parameters.StartRow];delete a[Microsoft.Office.WebExtension.Parameters.StartColumn];delete a[Microsoft.Office.WebExtension.Parameters.RowCount];delete a[Microsoft.Office.WebExtension.Parameters.ColumnCount]}if(a[Microsoft.Office.WebExtension.Parameters.CoercionType]!=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(b.type)&&(a[Microsoft.Office.WebExtension.Parameters.StartRow]||a[Microsoft.Office.WebExtension.Parameters.StartColumn]||a[Microsoft.Office.WebExtension.Parameters.RowCount]||a[Microsoft.Office.WebExtension.Parameters.ColumnCount]))throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;return a},privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:f}],onSucceeded:g});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:[b,c,d,"boolean"]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(a){return OSF.DDA.DataCoercion.determineCoercionType(a[Microsoft.Office.WebExtension.Parameters.Data])}}},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:{types:[c,b],defaultValue:a}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[c],defaultValue:a}},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:{types:[d],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:{types:[d],defaultValue:0}}],checkCallArgs:function(a,b){if(a[Microsoft.Office.WebExtension.Parameters.StartRow]==0&&a[Microsoft.Office.WebExtension.Parameters.StartColumn]==0){delete a[Microsoft.Office.WebExtension.Parameters.StartRow];delete a[Microsoft.Office.WebExtension.Parameters.StartColumn]}if(a[Microsoft.Office.WebExtension.Parameters.CoercionType]!=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(b.type)&&(a[Microsoft.Office.WebExtension.Parameters.StartRow]||a[Microsoft.Office.WebExtension.Parameters.StartColumn]))throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;return a},privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:f}]})})();Microsoft.Office.WebExtension.TableData=function(b,a){function c(a){if(a==null||a==undefined)return null;try{for(var b=OSF.DDA.DataCoercion.findArrayDimensionality(a,2);b<2;b++)a=[a];return a}catch(c){}}OSF.OUtil.defineEnumerableProperties(this,{headers:{"get":function(){return a},"set":function(b){a=c(b)}},rows:{"get":function(){return b},"set":function(a){b=a==null||OSF.OUtil.isArray(a)&&a.length==0?[]:c(a)}}});this.headers=a;this.rows=b};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureTableData=function(a){return new Microsoft.Office.WebExtension.TableData(a[OSF.DDA.TableDataProperties.TableRows],a[OSF.DDA.TableDataProperties.TableHeaders])};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{TableDataProperties:"TableDataProperties"});OSF.OUtil.augmentList(OSF.DDA.BindingProperties,{RowCount:"BindingRowCount",ColumnCount:"BindingColumnCount",HasHeaders:"HasHeaders"});OSF.DDA.TableDataProperties={TableRows:"TableRows",TableHeaders:"TableHeaders"};OSF.DDA.TableBinding=function(f,e,d,c,b){OSF.DDA.TableBinding.uber.constructor.call(this,f,e);OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.BindingType.Table},rowCount:{value:d?d:0},columnCount:{value:c?c:0},hasHeaders:{value:b?b:false}});var a=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[a.AddRowsAsync,a.AddColumnsAsync,a.DeleteAllDataValuesAsync])};OSF.OUtil.extend(OSF.DDA.TableBinding,OSF.DDA.Binding);OSF.DDA.AsyncMethodNames.addNames({AddRowsAsync:"addRowsAsync",AddColumnsAsync:"addColumnsAsync",DeleteAllDataValuesAsync:"deleteAllDataValuesAsync"});(function(){function a(a){return a.id}OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddRowsAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:a}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddColumnsAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:a}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:a}]})})();Microsoft.Office.WebExtension.CoercionType={Text:"text",Matrix:"matrix",Table:"table"};OSF.DDA.DataCoercion=function(){var a=null;return {findArrayDimensionality:function(c){if(OSF.OUtil.isArray(c)){for(var b=0,a=0;a<c.length;a++)b=Math.max(b,OSF.DDA.DataCoercion.findArrayDimensionality(c[a]));return b+1}else return 0},getCoercionDefaultForBinding:function(a){switch(a){case Microsoft.Office.WebExtension.BindingType.Matrix:return Microsoft.Office.WebExtension.CoercionType.Matrix;case Microsoft.Office.WebExtension.BindingType.Table:return Microsoft.Office.WebExtension.CoercionType.Table;case Microsoft.Office.WebExtension.BindingType.Text:default:return Microsoft.Office.WebExtension.CoercionType.Text}},getBindingDefaultForCoercion:function(a){switch(a){case Microsoft.Office.WebExtension.CoercionType.Matrix:return Microsoft.Office.WebExtension.BindingType.Matrix;case Microsoft.Office.WebExtension.CoercionType.Table:return Microsoft.Office.WebExtension.BindingType.Table;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:return Microsoft.Office.WebExtension.BindingType.Text}},determineCoercionType:function(b){if(b==a||b==undefined)return a;var c=a,d=typeof b;if(b.rows!==undefined)c=Microsoft.Office.WebExtension.CoercionType.Table;else if(OSF.OUtil.isArray(b))c=Microsoft.Office.WebExtension.CoercionType.Matrix;else if(d=="string"||d=="number"||d=="boolean"||OSF.OUtil.isDate(b))c=Microsoft.Office.WebExtension.CoercionType.Text;else throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;return c},coerceData:function(b,c,a){a=a||OSF.DDA.DataCoercion.determineCoercionType(b);if(a&&a!=c){OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin);b=OSF.DDA.DataCoercion._coerceDataFromTable(c,OSF.DDA.DataCoercion._coerceDataToTable(b,a));OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd)}return b},_matrixToText:function(a){if(a.length==1&&a[0].length==1)return ""+a[0][0];for(var b="",c=0;c<a.length;c++)b+=a[c].join("\t")+"\n";return b.substring(0,b.length-1)},_textToMatrix:function(c){for(var a=c.split("\n"),b=0;b<a.length;b++)a[b]=a[b].split("\t");return a},_tableToText:function(c){var b="";if(c.headers!=a)b=OSF.DDA.DataCoercion._matrixToText([c.headers])+"\n";var d=OSF.DDA.DataCoercion._matrixToText(c.rows);if(d=="")b=b.substring(0,b.length-1);return b+d},_tableToMatrix:function(b){var c=b.rows;b.headers!=a&&c.unshift(b.headers);return c},_coerceDataFromTable:function(d,c){var b;switch(d){case Microsoft.Office.WebExtension.CoercionType.Table:b=c;break;case Microsoft.Office.WebExtension.CoercionType.Matrix:b=OSF.DDA.DataCoercion._tableToMatrix(c);break;case Microsoft.Office.WebExtension.CoercionType.SlideRange:b=a;if(OSF.DDA.OMFactory.manufactureSlideRange)b=OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(c));if(b==a)b=OSF.DDA.DataCoercion._tableToText(c);break;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:b=OSF.DDA.DataCoercion._tableToText(c)}return b},_coerceDataToTable:function(b,c){if(c==undefined)c=OSF.DDA.DataCoercion.determineCoercionType(b);var a;switch(c){case Microsoft.Office.WebExtension.CoercionType.Table:a=b;break;case Microsoft.Office.WebExtension.CoercionType.Matrix:a=new Microsoft.Office.WebExtension.TableData(b);break;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:a=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(b))}return a}}}();OSF.DDA.FilePropertiesDescriptor={Url:"Url"};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{FilePropertiesDescriptor:"FilePropertiesDescriptor"});Microsoft.Office.WebExtension.FileProperties=function(a){OSF.OUtil.defineEnumerableProperties(this,{url:{value:a[OSF.DDA.FilePropertiesDescriptor.Url]}})};OSF.DDA.AsyncMethodNames.addNames({GetFilePropertiesAsync:"getFilePropertiesAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,fromHost:[{name:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,value:0}],requiredArguments:[],supportedOptions:[],onSucceeded:function(a){return new Microsoft.Office.WebExtension.FileProperties(a)}});Microsoft.Office.WebExtension.GoToType={Binding:"binding",NamedItem:"namedItem",Slide:"slide",Index:"index"};Microsoft.Office.WebExtension.SelectionMode={Default:"default",Selected:"selected",None:"none"};Microsoft.Office.WebExtension.Index={First:"first",Last:"last",Next:"next",Previous:"previous"};OSF.DDA.AsyncMethodNames.addNames({GoToByIdAsync:"goToByIdAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GoToByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:["string","number"]},{name:Microsoft.Office.WebExtension.Parameters.GoToType,"enum":Microsoft.Office.WebExtension.GoToType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.SelectionMode,value:{"enum":Microsoft.Office.WebExtension.SelectionMode,defaultValue:Microsoft.Office.WebExtension.SelectionMode.Default}}]});Microsoft.Office.WebExtension.EventType={};OSF.EventDispatch=function(c){var b=this;b._eventHandlers={};b._objectEventHandlers={};b._queuedEventsArgs={};if(c!=null)for(var d=0;d<c.length;d++){var a=c[d],e=a=="objectDeleted"||a=="objectSelectionChanged"||a=="objectDataChanged"||a=="contentControlAdded";if(!e)b._eventHandlers[a]=[];else b._objectEventHandlers[a]={};b._queuedEventsArgs[a]=[]}};OSF.EventDispatch.prototype={getSupportedEvents:function(){var a=[];for(var b in this._eventHandlers)a.push(b);for(var b in this._objectEventHandlers)a.push(b);return a},supportsEvent:function(b){for(var a in this._eventHandlers)if(b==a)return true;for(var a in this._objectEventHandlers)if(b==a)return true;return false},hasEventHandler:function(c,d){var a=this._eventHandlers[c];if(a&&a.length>0)for(var b=0;b<a.length;b++)if(a[b]===d)return true;return false},hasObjectEventHandler:function(d,e,f){var c=this._objectEventHandlers[d];if(c!=null)for(var a=c[e],b=0;a!=null&&b<a.length;b++)if(a[b]===f)return true;return false},addEventHandler:function(b,a){if(typeof a!="function")return false;var c=this._eventHandlers[b];if(c&&!this.hasEventHandler(b,a)){c.push(a);return true}else return false},addObjectEventHandler:function(d,b,c){if(typeof c!="function")return false;var a=this._objectEventHandlers[d];if(a&&!this.hasObjectEventHandler(d,b,c)){if(a[b]==null)a[b]=[];a[b].push(c);return true}return false},addEventHandlerAndFireQueuedEvent:function(a,e){var d=this._eventHandlers[a],c=d.length==0,b=this.addEventHandler(a,e);c&&b&&this.fireQueuedEvent(a);return b},removeEventHandler:function(c,d){var a=this._eventHandlers[c];if(a&&a.length>0)for(var b=0;b<a.length;b++)if(a[b]===d){a.splice(b,1);return true}return false},removeObjectEventHandler:function(d,e,f){var c=this._objectEventHandlers[d];if(c!=null)for(var a=c[e],b=0;a!=null&&b<a.length;b++)if(a[b]===f){a.splice(b,1);return true}return false},clearEventHandlers:function(a){if(typeof this._eventHandlers[a]!="undefined"&&this._eventHandlers[a].length>0){this._eventHandlers[a]=[];return true}return false},clearObjectEventHandlers:function(a,b){if(this._objectEventHandlers[a]!=null&&this._objectEventHandlers[a][b]!=null){this._objectEventHandlers[a][b]=[];return true}return false},getEventHandlerCount:function(a){return this._eventHandlers[a]!=undefined?this._eventHandlers[a].length:-1},getObjectEventHandlerCount:function(a,b){if(this._objectEventHandlers[a]==null||this._objectEventHandlers[a][b]==null)return 0;return this._objectEventHandlers[a][b].length},fireEvent:function(a){if(a.type==undefined)return false;var b=a.type;if(b&&this._eventHandlers[b]){for(var d=this._eventHandlers[b],c=0;c<d.length;c++)d[c](a);return true}else return false},fireObjectEvent:function(f,a){if(a.type==undefined)return false;var b=a.type;if(b&&this._objectEventHandlers[b]){var e=this._objectEventHandlers[b],c=e[f];if(c!=null){for(var d=0;d<c.length;d++)c[d](a);return true}}return false},fireOrQueueEvent:function(c){var b=this,a=c.type;if(a&&b._eventHandlers[a]){var d=b._eventHandlers[a],e=b._queuedEventsArgs[a];if(d.length==0)e.push(c);else b.fireEvent(c);return true}else return false},fireQueuedEvent:function(a){if(a&&this._eventHandlers[a]){var b=this._eventHandlers[a],c=this._queuedEventsArgs[a];if(b.length>0){var d=b[0];while(c.length>0){var e=c.shift();d(e)}return true}}return false},clearQueuedEvent:function(a){if(a&&this._eventHandlers[a]){var b=this._queuedEventsArgs[a];if(b)this._queuedEventsArgs[a]=[]}}};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureEventArgs=function(d,c,b){var h="hostPlatform",g="outlook",f="hostType",e=this,a;switch(d){case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:a=new OSF.DDA.DocumentSelectionChangedEventArgs(c);break;case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:a=new OSF.DDA.BindingSelectionChangedEventArgs(e.manufactureBinding(b,c.document),b[OSF.DDA.PropertyDescriptors.Subset]);break;case Microsoft.Office.WebExtension.EventType.BindingDataChanged:a=new OSF.DDA.BindingDataChangedEventArgs(e.manufactureBinding(b,c.document));break;case Microsoft.Office.WebExtension.EventType.SettingsChanged:a=new OSF.DDA.SettingsChangedEventArgs(c);break;case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:a=new OSF.DDA.ActiveViewChangedEventArgs(b);break;case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:a=new OSF.DDA.Theming.OfficeThemeChangedEventArgs(b);break;case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:a=new OSF.DDA.Theming.DocumentThemeChangedEventArgs(b);break;case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:a=OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(b);break;case Microsoft.Office.WebExtension.EventType.ObjectDeleted:case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:case Microsoft.Office.WebExtension.EventType.ContentControlAdded:a=new OSF.DDA.ObjectEventArgs(d,b[Microsoft.Office.WebExtension.Parameters.Id]);break;case Microsoft.Office.WebExtension.EventType.RichApiMessage:a=new OSF.DDA.RichApiMessageEventArgs(d,b);break;case Microsoft.Office.WebExtension.EventType.DataNodeInserted:a=new OSF.DDA.NodeInsertedEventArgs(e.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NewNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:a=new OSF.DDA.NodeReplacedEventArgs(e.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.OldNode]),e.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NewNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:a=new OSF.DDA.NodeDeletedEventArgs(e.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.OldNode]),e.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NextSiblingNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:a=new OSF.DDA.TaskSelectionChangedEventArgs(c);break;case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:a=new OSF.DDA.ResourceSelectionChangedEventArgs(c);break;case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:a=new OSF.DDA.ViewSelectionChangedEventArgs(c);break;case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:a=new OSF.DDA.DialogEventArgs(b);break;case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:a=new OSF.DDA.DialogParentEventArgs(b);break;case Microsoft.Office.WebExtension.EventType.ItemChanged:if(OSF._OfficeAppFactory.getHostInfo()[f]==g){a=new OSF.DDA.OlkItemSelectedChangedEventArgs(b);c.initialize(a["initialData"]);(OSF._OfficeAppFactory.getHostInfo()[h]=="win32"||OSF._OfficeAppFactory.getHostInfo()[h]=="mac")&&c.setCurrentItemNumber(a["itemNumber"].itemNumber)}else throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,d));break;case Microsoft.Office.WebExtension.EventType.RecipientsChanged:if(OSF._OfficeAppFactory.getHostInfo()[f]==g)a=new OSF.DDA.OlkRecipientsChangedEventArgs(b);else throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,d));break;case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:if(OSF._OfficeAppFactory.getHostInfo()[f]==g)a=new OSF.DDA.OlkAppointmentTimeChangedEventArgs(b);else throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,d));break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,d))}return a};OSF.DDA.AsyncMethodNames.addNames({AddHandlerAsync:"addHandlerAsync",RemoveHandlerAsync:"removeHandlerAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddHandlerAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(b,c,a){return a.supportsEvent(b)}},{name:Microsoft.Office.WebExtension.Parameters.Handler,types:["function"]}],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(b,c,a){return a.supportsEvent(b)}}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Handler,value:{types:["function","object"],defaultValue:null}}],privateStateCallbacks:[]});OSF.DialogShownStatus={hasDialogShown:false,isWindowDialog:false};OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{DialogMessageReceivedEvent:"DialogMessageReceivedEvent"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{DialogMessageReceived:"dialogMessageReceived",DialogEventReceived:"dialogEventReceived"});OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{MessageType:"messageType",MessageContent:"messageContent"});OSF.DDA.DialogEventType={};OSF.OUtil.augmentList(OSF.DDA.DialogEventType,{DialogClosed:"dialogClosed",NavigationFailed:"naviationFailed"});OSF.DDA.AsyncMethodNames.addNames({DisplayDialogAsync:"displayDialogAsync",CloseAsync:"close"});OSF.DDA.SyncMethodNames.addNames({MessageParent:"messageParent",AddMessageHandler:"addEventHandler",SendMessage:"sendMessage"});OSF.DDA.UI.ParentUI=function(){var c=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived,Microsoft.Office.WebExtension.EventType.DialogEventReceived,Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived]),b=OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName,a=this;!a[b]&&OSF.OUtil.defineEnumerableProperty(a,b,{value:function(){var b=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];b(arguments,c,a)}});OSF.OUtil.finalizeProperties(this)};OSF.DDA.UI.ChildUI=function(d){var b=OSF.DDA.SyncMethodNames.MessageParent.displayName,a=this;!a[b]&&OSF.OUtil.defineEnumerableProperty(a,b,{value:function(){var b=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];return b(arguments,a)}});var c=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;!a[c]&&typeof OSF.DialogParentMessageEventDispatch!="undefined"&&OSF.DDA.DispIdHost.addEventSupport(a,OSF.DialogParentMessageEventDispatch,d);OSF.OUtil.finalizeProperties(this)};OSF.DialogHandler=function(){};OSF.DDA.DialogEventArgs=function(a){if(a[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogMessageReceived)OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DialogMessageReceived},message:{value:a[OSF.DDA.PropertyDescriptors.MessageContent]}});else OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DialogEventReceived},error:{value:a[OSF.DDA.PropertyDescriptors.MessageType]}})};OSF.DDA.DialogParentEventArgs=function(a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived},message:{value:a[OSF.DDA.PropertyDescriptors.MessageContent]}})};OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.DisplayDialogAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Url,types:["string"]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Width,value:{types:["number"],defaultValue:99}},{name:Microsoft.Office.WebExtension.Parameters.Height,value:{types:["number"],defaultValue:99}},{name:Microsoft.Office.WebExtension.Parameters.RequireHTTPs,value:{types:["boolean"],defaultValue:true}},{name:Microsoft.Office.WebExtension.Parameters.DisplayInIframe,value:{types:["boolean"],defaultValue:false}},{name:Microsoft.Office.WebExtension.Parameters.HideTitle,value:{types:["boolean"],defaultValue:false}},{name:Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,value:{types:["boolean"],defaultValue:false}}],privateStateCallbacks:[],onSucceeded:function(c){var g=c[Microsoft.Office.WebExtension.Parameters.Id],b=c[Microsoft.Office.WebExtension.Parameters.Data],a=new OSF.DialogHandler,d=OSF.DDA.AsyncMethodNames.CloseAsync.displayName;OSF.OUtil.defineEnumerableProperty(a,d,{value:function(){var c=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];c(arguments,g,b,a)}});var f=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;OSF.OUtil.defineEnumerableProperty(a,f,{value:function(){var d=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id],c=d.verifyAndExtractCall(arguments,a,b),e=c[Microsoft.Office.WebExtension.Parameters.EventType],f=c[Microsoft.Office.WebExtension.Parameters.Handler];return b.addEventHandlerAndFireQueuedEvent(e,f)}});var e=OSF.DDA.SyncMethodNames.SendMessage.displayName;OSF.OUtil.defineEnumerableProperty(a,e,{value:function(){var c=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];return c(arguments,b,a)}});return a},checkCallArgs:function(a){if(a[Microsoft.Office.WebExtension.Parameters.Width]<=0)a[Microsoft.Office.WebExtension.Parameters.Width]=1;if(!a[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels]&&a[Microsoft.Office.WebExtension.Parameters.Width]>100)a[Microsoft.Office.WebExtension.Parameters.Width]=99;if(a[Microsoft.Office.WebExtension.Parameters.Height]<=0)a[Microsoft.Office.WebExtension.Parameters.Height]=1;if(!a[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels]&&a[Microsoft.Office.WebExtension.Parameters.Height]>100)a[Microsoft.Office.WebExtension.Parameters.Height]=99;if(!a[Microsoft.Office.WebExtension.Parameters.RequireHTTPs])a[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]=true;return a}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.CloseAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.SyncMethodCalls.define({method:OSF.DDA.SyncMethodNames.MessageParent,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.MessageToParent,types:["string","number","boolean"]}],supportedOptions:[]});OSF.DDA.SyncMethodCalls.define({method:OSF.DDA.SyncMethodNames.AddMessageHandler,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(b,c,a){return a.supportsEvent(b)}},{name:Microsoft.Office.WebExtension.Parameters.Handler,types:["function"]}],supportedOptions:[]});OSF.DDA.SyncMethodCalls.define({method:OSF.DDA.SyncMethodNames.SendMessage,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.MessageContent,types:["string"]}],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.SafeArray.Delegate.openDialog=function(a){try{a.onCalling&&a.onCalling();var c=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true,a);OSF.ClientHostController.openDialog(a.dispId,a.targetId,function(c,b){a.onEvent&&a.onEvent(b);OSF.AppTelemetry&&OSF.AppTelemetry.onEventDone(a.dispId)},c)}catch(b){OSF.DDA.SafeArray.Delegate._onException(b,a)}};OSF.DDA.SafeArray.Delegate.closeDialog=function(a){a.onCalling&&a.onCalling();var c=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false,a);try{OSF.ClientHostController.closeDialog(a.dispId,a.targetId,c)}catch(b){OSF.DDA.SafeArray.Delegate._onException(b,a)}};OSF.DDA.SafeArray.Delegate.messageParent=function(a){try{a.onCalling&&a.onCalling();var d=(new Date).getTime(),b=OSF.ClientHostController.messageParent(a.hostCallArgs);a.onReceiving&&a.onReceiving();OSF.AppTelemetry&&OSF.AppTelemetry.onMethodDone(a.dispId,a.hostCallArgs,Math.abs((new Date).getTime()-d),b);return b}catch(c){return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(c)}};OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}],isComplexType:true});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,fromHost:[{name:OSF.DDA.PropertyDescriptors.MessageType,value:0},{name:OSF.DDA.PropertyDescriptors.MessageContent,value:1}],isComplexType:true});OSF.DDA.SafeArray.Delegate.sendMessage=function(a){try{a.onCalling&&a.onCalling();var d=(new Date).getTime(),c=OSF.ClientHostController.sendMessage(a.hostCallArgs);a.onReceiving&&a.onReceiving();return c}catch(b){return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(b)}};OSF.OUtil.redefineList(Microsoft.Office.WebExtension.CoercionType,{Text:"text"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.ValueFormat,{Unformatted:"unformatted"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.FilterType,{All:"all"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.GoToType,{Slide:"slide",Index:"index"});delete Microsoft.Office.WebExtension.BindingType;delete Microsoft.Office.WebExtension.select;OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.Text,value:0},{name:Microsoft.Office.WebExtension.CoercionType.Matrix,value:1},{name:Microsoft.Office.WebExtension.CoercionType.Table,value:2}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{SlideRange:"slideRange"});OSF.DDA.SlideProperties={Id:0,Title:1,Index:2};OSF.DDA.Slide=function(c,b,a){OSF.OUtil.defineEnumerableProperties(this,{id:{value:c},title:{value:b},index:{value:a}})};OSF.DDA.SlideRange=function(a){OSF.OUtil.defineEnumerableProperties(this,{slides:{value:a}})};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureSlideRange=function(h){var c=null,a=c;if(JSON)a=JSON.parse(h);else a=Sys.Serialization.JavaScriptSerializer.deserialize(h);if(a==c)return c;var e=0;for(var k in OSF.DDA.SlideProperties)if(OSF.DDA.SlideProperties.hasOwnProperty(k))e++;for(var f=[],d=true,b=0;b<a.length&&d;b++){d=false;if(a[b].length==e){var i=parseInt(a[b][OSF.DDA.SlideProperties.Id]),j=a[b][OSF.DDA.SlideProperties.Title],g=parseInt(a[b][OSF.DDA.SlideProperties.Index]);if(!isNaN(i)&&!isNaN(g)){d=true;f.push(new OSF.DDA.Slide(i,j,g))}}}if(!d)return c;return new OSF.DDA.SlideRange(f)};OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.SlideRange,value:7}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{Html:"html"});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.Html,value:3}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{Ooxml:"ooxml"});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.Ooxml,value:4}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{OoxmlPackage:"ooxmlPackage"});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.OoxmlPackage,value:5}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{PdfFile:"pdf"});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.PdfFile,value:6}]});Microsoft.Office.WebExtension.FileType={Text:"text",Compressed:"compressed",Pdf:"pdf"};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{FileProperties:"FileProperties",FileSliceProperties:"FileSliceProperties"});OSF.DDA.FileProperties={Handle:"FileHandle",FileSize:"FileSize",SliceSize:Microsoft.Office.WebExtension.Parameters.SliceSize};OSF.DDA.File=function(e,c,b){OSF.OUtil.defineEnumerableProperties(this,{size:{value:c},sliceCount:{value:Math.ceil(c/b)}});var a={};a[OSF.DDA.FileProperties.Handle]=e;a[OSF.DDA.FileProperties.SliceSize]=b;var d=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[d.GetDocumentCopyChunkAsync,d.ReleaseDocumentCopyAsync],a)};OSF.DDA.FileSliceOffset="fileSliceoffset";OSF.DDA.AsyncMethodNames.addNames({GetDocumentCopyAsync:"getFileAsync",GetDocumentCopyChunkAsync:"getSliceAsync",ReleaseDocumentCopyAsync:"closeAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.FileType,"enum":Microsoft.Office.WebExtension.FileType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.SliceSize,value:{types:["number"],defaultValue:4*1024*1024}}],checkCallArgs:function(b){var a=b[Microsoft.Office.WebExtension.Parameters.SliceSize];if(a<=0||a>4*1024*1024)throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;return b},onSucceeded:function(a,c,b){return new OSF.DDA.File(a[OSF.DDA.FileProperties.Handle],a[OSF.DDA.FileProperties.FileSize],b[Microsoft.Office.WebExtension.Parameters.SliceSize])}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.SliceIndex,types:["number"]}],privateStateCallbacks:[{name:OSF.DDA.FileProperties.Handle,value:function(b,a){return a[OSF.DDA.FileProperties.Handle]}},{name:OSF.DDA.FileProperties.SliceSize,value:function(b,a){return a[OSF.DDA.FileProperties.SliceSize]}}],checkCallArgs:function(a,d,c){var b=a[Microsoft.Office.WebExtension.Parameters.SliceIndex];if(b<0||b>=d.sliceCount)throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;a[OSF.DDA.FileSliceOffset]=parseInt((b*c[OSF.DDA.FileProperties.SliceSize]).toString());return a},onSucceeded:function(a,d,c){var b={};OSF.OUtil.defineEnumerableProperties(b,{data:{value:a[Microsoft.Office.WebExtension.Parameters.Data]},index:{value:c[Microsoft.Office.WebExtension.Parameters.SliceIndex]},size:{value:a[OSF.DDA.FileProperties.SliceSize]}});return b}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,privateStateCallbacks:[{name:OSF.DDA.FileProperties.Handle,value:function(b,a){return a[OSF.DDA.FileProperties.Handle]}}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FileProperties,fromHost:[{name:OSF.DDA.FileProperties.Handle,value:0},{name:OSF.DDA.FileProperties.FileSize,value:1}],isComplexType:true});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FileSliceProperties,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:0},{name:OSF.DDA.FileProperties.SliceSize,value:1}],isComplexType:true});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.FileType,toHost:[{name:Microsoft.Office.WebExtension.FileType.Text,value:0},{name:Microsoft.Office.WebExtension.FileType.Compressed,value:5},{name:Microsoft.Office.WebExtension.FileType.Pdf,value:6}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.FileType,value:0}],fromHost:[{name:OSF.DDA.PropertyDescriptors.FileProperties,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,toHost:[{name:OSF.DDA.FileProperties.Handle,value:0},{name:OSF.DDA.FileSliceOffset,value:1},{name:OSF.DDA.FileProperties.SliceSize,value:2}],fromHost:[{name:OSF.DDA.PropertyDescriptors.FileSliceProperties,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,toHost:[{name:OSF.DDA.FileProperties.Handle,value:0}]});OSF.DDA.AsyncMethodNames.addNames({GetSelectedDataAsync:"getSelectedDataAsync",SetSelectedDataAsync:"setSelectedDataAsync"});(function(){var c=false,b="boolean",a="number";function d(b,d,c){var a=b[Microsoft.Office.WebExtension.Parameters.Data];if(OSF.DDA.TableDataProperties&&a&&(a[OSF.DDA.TableDataProperties.TableRows]!=undefined||a[OSF.DDA.TableDataProperties.TableHeaders]!=undefined))a=OSF.DDA.OMFactory.manufactureTableData(a);a=OSF.DDA.DataCoercion.coerceData(a,c[Microsoft.Office.WebExtension.Parameters.CoercionType]);return a==undefined?null:a}OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,"enum":Microsoft.Office.WebExtension.CoercionType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:{"enum":Microsoft.Office.WebExtension.ValueFormat,defaultValue:Microsoft.Office.WebExtension.ValueFormat.Unformatted}},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:{"enum":Microsoft.Office.WebExtension.FilterType,defaultValue:Microsoft.Office.WebExtension.FilterType.All}}],privateStateCallbacks:[],onSucceeded:d});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["string","object",a,b]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(a){return OSF.DDA.DataCoercion.determineCoercionType(a[Microsoft.Office.WebExtension.Parameters.Data])}}},{name:Microsoft.Office.WebExtension.Parameters.ImageLeft,value:{types:[a,b],defaultValue:c}},{name:Microsoft.Office.WebExtension.Parameters.ImageTop,value:{types:[a,b],defaultValue:c}},{name:Microsoft.Office.WebExtension.Parameters.ImageWidth,value:{types:[a,b],defaultValue:c}},{name:Microsoft.Office.WebExtension.Parameters.ImageHeight,value:{types:[a,b],defaultValue:c}}],privateStateCallbacks:[]})})();OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:0},{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:1},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:2}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:0},{name:Microsoft.Office.WebExtension.Parameters.Data,value:1},{name:Microsoft.Office.WebExtension.Parameters.ImageLeft,value:2},{name:Microsoft.Office.WebExtension.Parameters.ImageTop,value:3},{name:Microsoft.Office.WebExtension.Parameters.ImageWidth,value:4},{name:Microsoft.Office.WebExtension.Parameters.ImageHeight,value:5}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.GoToType,toHost:[{name:Microsoft.Office.WebExtension.GoToType.Binding,value:0},{name:Microsoft.Office.WebExtension.GoToType.NamedItem,value:1},{name:Microsoft.Office.WebExtension.GoToType.Slide,value:2},{name:Microsoft.Office.WebExtension.GoToType.Index,value:3}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.SelectionMode,toHost:[{name:Microsoft.Office.WebExtension.SelectionMode.Default,value:0},{name:Microsoft.Office.WebExtension.SelectionMode.Selected,value:1},{name:Microsoft.Office.WebExtension.SelectionMode.None,value:2}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidNavigateToMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0},{name:Microsoft.Office.WebExtension.Parameters.GoToType,value:1},{name:Microsoft.Office.WebExtension.Parameters.SelectionMode,value:2}]});OSF.DDA.SettingsManager={SerializedSettings:"serializedSettings",RefreshingSettings:"refreshingSettings",DateJSONPrefix:"Date(",DataJSONSuffix:")",serializeSettings:function(a){return OSF.OUtil.serializeSettings(a)},deserializeSettings:function(a){return OSF.OUtil.deserializeSettings(a)}};OSF.DDA.Settings=function(a){var b="name";a=a||{};var c=function(d){var b=OSF.OUtil.getSessionStorage();if(b){var a=OSF.DDA.SettingsManager.serializeSettings(d),c=JSON?JSON.stringify(a):Sys.Serialization.JavaScriptSerializer.serialize(a);b.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(),c)}};OSF.OUtil.defineEnumerableProperties(this,{"get":{value:function(e){var d=Function._validateParams(arguments,[{name:b,type:String,mayBeNull:false}]);if(d)throw d;var c=a[e];return typeof c==="undefined"?null:c}},"set":{value:function(f,e){var d=Function._validateParams(arguments,[{name:b,type:String,mayBeNull:false},{name:"value",mayBeNull:true}]);if(d)throw d;a[f]=e;c(a)}},remove:{value:function(e){var d=Function._validateParams(arguments,[{name:b,type:String,mayBeNull:false}]);if(d)throw d;delete a[e];c(a)}}});OSF.DDA.DispIdHost.addAsyncMethods(this,[OSF.DDA.AsyncMethodNames.SaveAsync],a)};OSF.DDA.RefreshableSettings=function(a){OSF.DDA.RefreshableSettings.uber.constructor.call(this,a);OSF.DDA.DispIdHost.addAsyncMethods(this,[OSF.DDA.AsyncMethodNames.RefreshAsync],a);OSF.DDA.DispIdHost.addEventSupport(this,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]))};OSF.OUtil.extend(OSF.DDA.RefreshableSettings,OSF.DDA.Settings);OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{SettingsChanged:"settingsChanged"});OSF.DDA.SettingsChangedEventArgs=function(a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.SettingsChanged},settings:{value:a}})};OSF.DDA.AsyncMethodNames.addNames({RefreshAsync:"refreshAsync",SaveAsync:"saveAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.RefreshAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.SettingsManager.RefreshingSettings,value:function(b,a){return a}}],onSucceeded:function(d,a,e){var f=d[OSF.DDA.SettingsManager.SerializedSettings],c=OSF.DDA.SettingsManager.deserializeSettings(f),g=e[OSF.DDA.SettingsManager.RefreshingSettings];for(var b in g)a.remove(b);for(var b in c)a.set(b,c[b]);return a}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SaveAsync,requiredArguments:[],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,value:{types:["boolean"],defaultValue:true}}],privateStateCallbacks:[{name:OSF.DDA.SettingsManager.SerializedSettings,value:function(b,a){return OSF.DDA.SettingsManager.serializeSettings(a)}}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidLoadSettingsMethod,fromHost:[{name:OSF.DDA.SettingsManager.SerializedSettings,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSaveSettingsMethod,toHost:[{name:OSF.DDA.SettingsManager.SerializedSettings,value:OSF.DDA.SettingsManager.SerializedSettings},{name:Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,value:Microsoft.Office.WebExtension.Parameters.OverwriteIfStale}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidSettingsChangedEvent});OSF.DDA.AsyncMethodNames.addNames({GetOfficeThemeAsync:"getOfficeThemeAsync",GetDocumentThemeAsync:"getDocumentThemeAsync"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{OfficeThemeChanged:"officeThemeChanged",DocumentThemeChanged:"documentThemeChanged"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.Parameters,{DocumentTheme:"documentTheme",OfficeTheme:"officeTheme"});OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{DocumentThemeChangedEvent:"DocumentThemeChangedEvent",OfficeThemeChangedEvent:"OfficeThemeChangedEvent"});OSF.OUtil.setNamespace("Theming",OSF.DDA);OSF.DDA.Theming.OfficeThemeEnum={PrimaryFontColor:"primaryFontColor",PrimaryBackgroundColor:"primaryBackgroundColor",SecondaryFontColor:"secondaryFontColor",SecondaryBackgroundColor:"secondaryBackgroundColor"};OSF.DDA.Theming.DocumentThemeEnum={PrimaryFontColor:"primaryFontColor",PrimaryBackgroundColor:"primaryBackgroundColor",SecondaryFontColor:"secondaryFontColor",SecondaryBackgroundColor:"secondaryBackgroundColor",Accent1:"accent1",Accent2:"accent2",Accent3:"accent3",Accent4:"accent4",Accent5:"accent5",Accent6:"accent6",Hyperlink:"hyperlink",FollowedHyperlink:"followedHyperlink",HeaderLatinFont:"headerLatinFont",HeaderEastAsianFont:"headerEastAsianFont",HeaderScriptFont:"headerScriptFont",HeaderLocalizedFont:"headerLocalizedFont",BodyLatinFont:"bodyLatinFont",BodyEastAsianFont:"bodyEastAsianFont",BodyScriptFont:"bodyScriptFont",BodyLocalizedFont:"bodyLocalizedFont"};OSF.DDA.Theming.ConvertToDocumentTheme=function(f){var b=false,a=true;for(var d=[{name:"primaryFontColor",needToConvertToHex:a},{name:"primaryBackgroundColor",needToConvertToHex:a},{name:"secondaryFontColor",needToConvertToHex:a},{name:"secondaryBackgroundColor",needToConvertToHex:a},{name:"accent1",needToConvertToHex:a},{name:"accent2",needToConvertToHex:a},{name:"accent3",needToConvertToHex:a},{name:"accent4",needToConvertToHex:a},{name:"accent5",needToConvertToHex:a},{name:"accent6",needToConvertToHex:a},{name:"hyperlink",needToConvertToHex:a},{name:"followedHyperlink",needToConvertToHex:a},{name:"headerLatinFont",needToConvertToHex:b},{name:"headerEastAsianFont",needToConvertToHex:b},{name:"headerScriptFont",needToConvertToHex:b},{name:"headerLocalizedFont",needToConvertToHex:b},{name:"bodyLatinFont",needToConvertToHex:b},{name:"bodyEastAsianFont",needToConvertToHex:b},{name:"bodyScriptFont",needToConvertToHex:b},{name:"bodyLocalizedFont",needToConvertToHex:b}],e={},c=0;c<d.length;c++)if(d[c].needToConvertToHex)e[d[c].name]=OSF.OUtil.convertIntToCssHexColor(f[d[c].name]);else e[d[c].name]=f[d[c].name];return e};OSF.DDA.Theming.ConvertToOfficeTheme=function(a){var b={};for(var c in a)b[c]=OSF.OUtil.convertIntToCssHexColor(a[c]);return b};OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDocumentThemeAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[],onSucceeded:OSF.DDA.Theming.ConvertToDocumentTheme});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetOfficeThemeAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[],onSucceeded:OSF.DDA.Theming.ConvertToOfficeTheme});OSF.DDA.Theming.OfficeThemeChangedEventArgs=function(a){var b=OSF.DDA.Theming.ConvertToOfficeTheme(a);OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.OfficeThemeChanged},officeTheme:{value:b}})};OSF.DDA.Theming.DocumentThemeChangedEventArgs=function(a){var b=OSF.DDA.Theming.ConvertToDocumentTheme(a);OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DocumentThemeChanged},documentTheme:{value:b}})};var OSF_DDA_Theming_InternalThemeHandler=function(){var h="secondaryBackgroundColor",g="secondaryFontColor",c="background-color",f="primaryBackgroundColor",b="color",e="primaryFontColor",a=null;function d(){var b=this;b._pseudoDocumentObject=a;b._previousDocumentThemeData=a;b._previousOfficeThemeData=a;b._officeCss=a;b._asyncCallsCompleted=a;b._onAsyncCallsCompleted=a}d.prototype.InitializeAndChangeOnce=function(c){var a=this;a._officeCss=a._getOfficeThemesCss();if(!a._officeCss){c&&c();return}a._onAsyncCallsCompleted=c;a._pseudoDocumentObject={};var b=a._pseudoDocumentObject;OSF.DDA.DispIdHost.addAsyncMethods(b,[OSF.DDA.AsyncMethodNames.GetOfficeThemeAsync,OSF.DDA.AsyncMethodNames.GetDocumentThemeAsync]);OSF.DDA.DispIdHost.addEventSupport(b,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.OfficeThemeChanged,Microsoft.Office.WebExtension.EventType.DocumentThemeChanged]));a._asyncCallsCompleted={};a._asyncCallsCompleted[OSF.DDA.AsyncMethodNames.GetOfficeThemeAsync]=false;a._asyncCallsCompleted[OSF.DDA.AsyncMethodNames.GetDocumentThemeAsync]=false;a._getAndProcessThemeData(b.getDocumentThemeAsync,Function.createDelegate(a,a._processDocumentThemeData),OSF.DDA.AsyncMethodNames.GetDocumentThemeAsync);a._getAndProcessThemeData(b.getOfficeThemeAsync,Function.createDelegate(a,a._processOfficeThemeData),OSF.DDA.AsyncMethodNames.GetOfficeThemeAsync)};d.prototype._getOfficeThemesCss=function(){function b(){for(var d="officethemes.css",c=0;c<document.styleSheets.length;c++){var b=document.styleSheets[c];if(!b.disabled&&b.href&&d==b.href.substring(b.href.length-d.length,b.href.length).toLowerCase())if(!b.cssRules&&!b.rules)return a;else return b}}try{return b()}catch(c){return a}};d.prototype._changeCss=function(a,f,e){for(var g=a.cssRules?a.cssRules.length:a.rules.length,b=0;b<g;b++){var d;if(a.cssRules)d=a.cssRules[b];else d=a.rules[b];var c=d.selectorText;if(c&&c.toLowerCase()==f.toLowerCase())if(a.cssRules){a.deleteRule(b);a.insertRule(c+e,b)}else{a.removeRule(b);a.addRule(c,e,b)}}};d.prototype._changeDocumentThemeData=function(u){var d="font-family",i="border-color",r="accent6",q="accent5",p="accent4",o="accent3",n="accent2",m="accent1",l=this;for(var k=[{name:e,cssSelector:".office-docTheme-primary-fontColor",cssProperty:b},{name:f,cssSelector:".office-docTheme-primary-bgColor",cssProperty:c},{name:g,cssSelector:".office-docTheme-secondary-fontColor",cssProperty:b},{name:h,cssSelector:".office-docTheme-secondary-bgColor",cssProperty:c},{name:m,cssSelector:".office-contentAccent1-color",cssProperty:b},{name:n,cssSelector:".office-contentAccent2-color",cssProperty:b},{name:o,cssSelector:".office-contentAccent3-color",cssProperty:b},{name:p,cssSelector:".office-contentAccent4-color",cssProperty:b},{name:q,cssSelector:".office-contentAccent5-color",cssProperty:b},{name:r,cssSelector:".office-contentAccent6-color",cssProperty:b},{name:m,cssSelector:".office-contentAccent1-bgColor",cssProperty:c},{name:n,cssSelector:".office-contentAccent2-bgColor",cssProperty:c},{name:o,cssSelector:".office-contentAccent3-bgColor",cssProperty:c},{name:p,cssSelector:".office-contentAccent4-bgColor",cssProperty:c},{name:q,cssSelector:".office-contentAccent5-bgColor",cssProperty:c},{name:r,cssSelector:".office-contentAccent6-bgColor",cssProperty:c},{name:m,cssSelector:".office-contentAccent1-borderColor",cssProperty:i},{name:n,cssSelector:".office-contentAccent2-borderColor",cssProperty:i},{name:o,cssSelector:".office-contentAccent3-borderColor",cssProperty:i},{name:p,cssSelector:".office-contentAccent4-borderColor",cssProperty:i},{name:q,cssSelector:".office-contentAccent5-borderColor",cssProperty:i},{name:r,cssSelector:".office-contentAccent6-borderColor",cssProperty:i},{name:"hyperlink",cssSelector:".office-a",cssProperty:b},{name:"followedHyperlink",cssSelector:".office-a:visited",cssProperty:b},{name:"headerLatinFont",cssSelector:".office-headerFont-latin",cssProperty:d},{name:"headerEastAsianFont",cssSelector:".office-headerFont-eastAsian",cssProperty:d},{name:"headerScriptFont",cssSelector:".office-headerFont-script",cssProperty:d},{name:"headerLocalizedFont",cssSelector:".office-headerFont-localized",cssProperty:d},{name:"bodyLatinFont",cssSelector:".office-bodyFont-latin",cssProperty:d},{name:"bodyEastAsianFont",cssSelector:".office-bodyFont-eastAsian",cssProperty:d},{name:"bodyScriptFont",cssSelector:".office-bodyFont-script",cssProperty:d},{name:"bodyLocalizedFont",cssSelector:".office-bodyFont-localized",cssProperty:d}],s=u.type=="documentThemeChanged"?u.documentTheme:u,j=0;j<k.length;j++)if(l._previousDocumentThemeData===a||l._previousDocumentThemeData[k[j].name]!=s[k[j].name])if(s[k[j].name]!=a&&s[k[j].name]!=""){var t=s[k[j].name];if(k[j].cssProperty===d)t='"'+t.replace(/"/g,'\\"')+'"';l._changeCss(l._officeCss,k[j].cssSelector,"{"+k[j].cssProperty+":"+t+";}")}else l._changeCss(l._officeCss,k[j].cssSelector,"{}");l._previousDocumentThemeData=s};d.prototype._changeOfficeThemeData=function(l){var j=this;for(var i=[{name:e,cssSelector:".office-officeTheme-primary-fontColor",cssProperty:b},{name:f,cssSelector:".office-officeTheme-primary-bgColor",cssProperty:c},{name:g,cssSelector:".office-officeTheme-secondary-fontColor",cssProperty:b},{name:h,cssSelector:".office-officeTheme-secondary-bgColor",cssProperty:c}],k=l.type=="officeThemeChanged"?l.officeTheme:l,d=0;d<i.length;d++)if(j._previousOfficeThemeData===a||j._previousOfficeThemeData[i[d].name]!=k[i[d].name])k[i[d].name]!==undefined&&j._changeCss(j._officeCss,i[d].cssSelector,"{"+i[d].cssProperty+":"+k[i[d].name]+";}");j._previousOfficeThemeData=k};d.prototype._getAndProcessThemeData=function(d,c,b){d(Function.createDelegate(this,function(e){var d=this;if(e.status=="succeeded"){var f=e.value;c(f)}if(d._areAllCallsCompleted(b)&&d._onAsyncCallsCompleted){d._onAsyncCallsCompleted();d._onAsyncCallsCompleted=a}}))};d.prototype._processOfficeThemeData=function(c){var b=this;b._changeOfficeThemeData(c);b._pseudoDocumentObject.addHandlerAsync(Microsoft.Office.WebExtension.EventType.OfficeThemeChanged,Function.createDelegate(b,b._changeOfficeThemeData),a)};d.prototype._processDocumentThemeData=function(c){var b=this;b._changeDocumentThemeData(c);b._pseudoDocumentObject.addHandlerAsync(Microsoft.Office.WebExtension.EventType.DocumentThemeChanged,Function.createDelegate(b,b._changeDocumentThemeData),a)};d.prototype._areAllCallsCompleted=function(b){var a;if(!(a=this._asyncCallsCompleted))return true;if(b&&a.hasOwnProperty(b))a[b]=true;for(var c in a){if(a.hasOwnProperty(c)&&a[c])continue;return false}return true};return d}();OSF.DDA.Theming.InternalThemeHandler=OSF_DDA_Theming_InternalThemeHandler;var parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;parameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDocumentThemeMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.DocumentTheme,value:parameterMap.self}]});parameterMap.define({type:OSF.DDA.MethodDispId.dispidGetOfficeThemeMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.OfficeTheme,value:parameterMap.self}]});parameterMap.define({type:OSF.DDA.EventDispId.dispidDocumentThemeChangedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,value:parameterMap.self}],isComplexType:true});parameterMap.define({type:OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,value:parameterMap.self}],isComplexType:true});var destKeys=OSF.DDA.Theming.DocumentThemeEnum;parameterMap.define({type:Microsoft.Office.WebExtension.Parameters.DocumentTheme,fromHost:[{name:destKeys.PrimaryBackgroundColor,value:0},{name:destKeys.PrimaryFontColor,value:1},{name:destKeys.SecondaryBackgroundColor,value:2},{name:destKeys.SecondaryFontColor,value:3},{name:destKeys.Accent1,value:4},{name:destKeys.Accent2,value:5},{name:destKeys.Accent3,value:6},{name:destKeys.Accent4,value:7},{name:destKeys.Accent5,value:8},{name:destKeys.Accent6,value:9},{name:destKeys.Hyperlink,value:10},{name:destKeys.FollowedHyperlink,value:11},{name:destKeys.HeaderLatinFont,value:12},{name:destKeys.HeaderEastAsianFont,value:13},{name:destKeys.HeaderScriptFont,value:14},{name:destKeys.HeaderLocalizedFont,value:15},{name:destKeys.BodyLatinFont,value:16},{name:destKeys.BodyEastAsianFont,value:17},{name:destKeys.BodyScriptFont,value:18},{name:destKeys.BodyLocalizedFont,value:19}],isComplexType:true});destKeys=OSF.DDA.Theming.OfficeThemeEnum;parameterMap.define({type:Microsoft.Office.WebExtension.Parameters.OfficeTheme,fromHost:[{name:destKeys.PrimaryFontColor,value:0},{name:destKeys.PrimaryBackgroundColor,value:1},{name:destKeys.SecondaryFontColor,value:2},{name:destKeys.SecondaryBackgroundColor,value:3}],isComplexType:true});parameterMap.define({type:OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.DocumentTheme,value:parameterMap.self}],isComplexType:true});parameterMap.define({type:OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.OfficeTheme,value:parameterMap.self}],isComplexType:true});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{DocumentSelectionChanged:"documentSelectionChanged"});OSF.DDA.DocumentSelectionChangedEventArgs=function(a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged},document:{value:a}})};OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ObjectDeleted:"objectDeleted"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ObjectSelectionChanged:"objectSelectionChanged"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ObjectDataChanged:"objectDataChanged"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ContentControlAdded:"contentControlAdded"});OSF.DDA.ObjectEventArgs=function(a,b){OSF.OUtil.defineEnumerableProperties(this,{type:{value:a},object:{value:b}})};OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidObjectDeletedEvent,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidObjectDataChangedEvent,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidContentControlAddedEvent,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,fromHost:[{name:OSF.DDA.FilePropertiesDescriptor.Url,value:0}],isComplexType:true});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,fromHost:[{name:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{ActiveViewChangedEvent:"ActiveViewChangedEvent"});Microsoft.Office.WebExtension.ActiveView={Read:"read",Edit:"edit"};OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ActiveViewChanged:"activeViewChanged"});OSF.DDA.AsyncMethodNames.addNames({GetActiveViewAsync:"getActiveViewAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetActiveViewAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(b){var a=b[Microsoft.Office.WebExtension.Parameters.ActiveView];return a==undefined?null:a}});OSF.DDA.ActiveViewChangedEventArgs=function(a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.ActiveViewChanged},activeView:{value:a.activeView}})};OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.ActiveView,fromHost:[{name:0,value:Microsoft.Office.WebExtension.ActiveView.Read},{name:1,value:Microsoft.Office.WebExtension.ActiveView.Edit}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetActiveViewMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.ActiveView,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDescriptors.ActiveViewChangedEvent,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.ActiveView,value:0}],isComplexType:true});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidActiveViewChangedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.ActiveViewChangedEvent,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{Image:"image"});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:Microsoft.Office.WebExtension.CoercionType.Image,value:8}]});var OfficeExt;(function(a){var b;(function(b){var e=function(){var f="object",g="string",d=null;function e(){var a=this,b=a;a._pseudoDocument=d;a._eventDispatch=d;a._processAppCommandInvocation=function(a){var c=b._verifyManifestCallback(a.callbackName);if(c.errorCode!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess){b._invokeAppCommandCompletedMethod(a.appCommandId,c.errorCode,"");return}var d=b._constructEventObjectForCallback(a);if(d)window.setTimeout(function(){c.callback(d)},0);else b._invokeAppCommandCompletedMethod(a.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,"")}}e.initializeOsfDda=function(){OSF.DDA.AsyncMethodNames.addNames({AppCommandInvocationCompletedAsync:"appCommandInvocationCompletedAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[g]},{name:Microsoft.Office.WebExtension.Parameters.Status,types:["number"]},{name:Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,types:[g]}]});OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{AppCommandInvokedEvent:"AppCommandInvokedEvent"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AppCommandInvoked:"appCommandInvoked"});OSF.OUtil.setNamespace("AppCommand",OSF.DDA);OSF.DDA.AppCommand.AppCommandInvokedEventArgs=a.AppCommand.AppCommandInvokedEventArgs};e.prototype.initializeAndChangeOnce=function(c){var a=this;b.registerDdaFacade();a._pseudoDocument={};OSF.DDA.DispIdHost.addAsyncMethods(a._pseudoDocument,[OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync]);a._eventDispatch=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.AppCommandInvoked]);var d=function(a){if(c)if(a.status=="succeeded")c(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);else c(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)};OSF.DDA.DispIdHost.addEventSupport(a._pseudoDocument,a._eventDispatch);a._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked,a._processAppCommandInvocation,d)};e.prototype._verifyManifestCallback=function(h){var a="function",g={callback:d,errorCode:OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback};h=h.trim();try{for(var b=h.split("."),c=window,e=0;e<b.length-1;e++)if(c[b[e]]&&(typeof c[b[e]]==f||typeof c[b[e]]==a))c=c[b[e]];else return g;var i=c[b[b.length-1]];if(typeof i!=a)return g}catch(j){return g}return {callback:i,errorCode:OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess}};e.prototype._invokeAppCommandCompletedMethod=function(a,b,c){this._pseudoDocument.appCommandInvocationCompletedAsync(a,b,c)};e.prototype._constructEventObjectForCallback=function(b){var f=this,a=new c;try{var e=JSON.parse(b.eventObjStr);this._translateEventObjectInternal(e,a);Object.defineProperty(a,"completed",{value:function(c){a.completedContext=c;var d=JSON.stringify(a);f._invokeAppCommandCompletedMethod(b.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess,d)},enumerable:true})}catch(g){a=d}return a};e.prototype._translateEventObjectInternal=function(e,c){for(var a in e){if(!e.hasOwnProperty(a))continue;var b=e[a];if(typeof b==f&&b!=d){OSF.OUtil.defineEnumerableProperty(c,a,{value:{}});this._translateEventObjectInternal(b,c[a])}else Object.defineProperty(c,a,{value:b,enumerable:true,writable:true})}};e.prototype._constructObjectByTemplate=function(c,j){var b={};if(!c||!j)return b;for(var a in c)if(c.hasOwnProperty(a)){b[a]=d;if(j[a]!=d){var h=c[a],i=j[a],e=typeof i;if(typeof h==f&&h!=d)b[a]=this._constructObjectByTemplate(h,i);else if(e=="number"||e==g||e=="boolean")b[a]=i}}return b};e.instance=function(){if(e._instance==d)e._instance=new e;return e._instance};e._instance=d;return e}();b.AppCommandManager=e;var d=function(){function a(b,c,d){var a=this;a.type=Microsoft.Office.WebExtension.EventType.AppCommandInvoked;a.appCommandId=b;a.callbackName=c;a.eventObjStr=d}a.create=function(c){return new a(c[b.AppCommandInvokedEventEnums.AppCommandId],c[b.AppCommandInvokedEventEnums.CallbackName],c[b.AppCommandInvokedEventEnums.EventObjStr])};return a}();b.AppCommandInvokedEventArgs=d;var c=function(){function a(){}return a}();b.AppCommandCallbackEventArgs=c;b.AppCommandInvokedEventEnums={AppCommandId:"appCommandId",CallbackName:"callbackName",EventObjStr:"eventObjStr"}})(b=a.AppCommand||(a.AppCommand={}))})(OfficeExt||(OfficeExt={}));OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();var OfficeExt;(function(a){var b;(function(c){function b(){if(OSF.DDA.SafeArray){var b=OSF.DDA.SafeArray.Delegate.ParameterMap;b.define({type:OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:0},{name:Microsoft.Office.WebExtension.Parameters.Status,value:1},{name:Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,value:2}]});b.define({type:OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.AppCommandInvokedEvent,value:b.self}],isComplexType:true});b.define({type:OSF.DDA.EventDescriptors.AppCommandInvokedEvent,fromHost:[{name:a.AppCommand.AppCommandInvokedEventEnums.AppCommandId,value:0},{name:a.AppCommand.AppCommandInvokedEventEnums.CallbackName,value:1},{name:a.AppCommand.AppCommandInvokedEventEnums.EventObjStr,value:2}],isComplexType:true})}}c.registerDdaFacade=b})(b=a.AppCommand||(a.AppCommand={}))})(OfficeExt||(OfficeExt={}));OSF.DDA.AsyncMethodNames.addNames({GetAccessTokenAsync:"getAccessTokenAsync"});OSF.DDA.Auth=function(){};OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,requiredArguments:[],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.ForceConsent,value:{types:["boolean"],defaultValue:false}},{name:Microsoft.Office.WebExtension.Parameters.ForceAddAccount,value:{types:["boolean"],defaultValue:false}},{name:Microsoft.Office.WebExtension.Parameters.AuthChallenge,value:{types:["string"],defaultValue:""}}],onSucceeded:function(a){var b=a[Microsoft.Office.WebExtension.Parameters.Data];return b}});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.ForceConsent,value:0},{name:Microsoft.Office.WebExtension.Parameters.ForceAddAccount,value:1},{name:Microsoft.Office.WebExtension.Parameters.AuthChallenge,value:2}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.AsyncMethodNames.addNames({ExecuteRichApiRequestAsync:"executeRichApiRequestAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[]});OSF.OUtil.setNamespace("RichApi",OSF.DDA);OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:0}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.SafeArray.Delegate.ParameterMap.self}]});OSF.DDA.AsyncMethodNames.addNames({OpenBrowserWindow:"openBrowserWindow"});OSF.DDA.OpenBrowser=function(){};OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.OpenBrowserWindow,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Url,types:["string"]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Reserved,value:{types:["number"],defaultValue:0}}],privateStateCallbacks:[]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidOpenBrowserWindow,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Reserved,value:0},{name:Microsoft.Office.WebExtension.Parameters.Url,value:1}]});OSF.DDA.AsyncMethodNames.addNames({CreateDocumentAsync:"createDocumentAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.CreateDocumentAsync,supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Base64,value:{types:["string"],defaultValue:""}}],privateStateCallbacks:[]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidCreateDocumentMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Base64,value:0}]});OSF.DDA.AsyncMethodNames.addNames({InsertFormAsync:"insertFormAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.InsertFormAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.FormId,types:["string"]}],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.SafeArray.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidInsertFormMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.FormId,value:0}]});OSF.DDA.PowerPointDocument=function(b,c){var a=this;OSF.DDA.PowerPointDocument.uber.constructor.call(a,b,c);OSF.DDA.DispIdHost.addAsyncMethods(a,[OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,OSF.DDA.AsyncMethodNames.GetActiveViewAsync,OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,OSF.DDA.AsyncMethodNames.GoToByIdAsync,OSF.DDA.AsyncMethodNames.InsertFormAsync]);OSF.DDA.DispIdHost.addEventSupport(a,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged,Microsoft.Office.WebExtension.EventType.ActiveViewChanged]));OSF.OUtil.finalizeProperties(a)};OSF.OUtil.extend(OSF.DDA.PowerPointDocument,OSF.DDA.Document);OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize=function(){var a=OfficeExt.AppCommand.AppCommandManager.instance();a.initializeAndChangeOnce()};OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM=function(a,b){OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);a.doc=new OSF.DDA.PowerPointDocument(a,this._initializeSettings(true));OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi,[OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);b()};OSF.InitializationHelper.prototype.prepareApiSurface=function(a){var c=new OSF.DDA.License(a.get_eToken());if(a.get_isDialog())a.ui=new OSF.DDA.UI.ChildUI;else a.ui=new OSF.DDA.UI.ParentUI;OSF.DDA.OpenBrowser&&OSF.DDA.DispIdHost.addAsyncMethods(a.ui,[OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);if(OSF.DDA.Auth){a.auth=new OSF.DDA.Auth;OSF.DDA.DispIdHost.addAsyncMethods(a.auth,[OSF.DDA.AsyncMethodNames.GetAccessTokenAsync])}a.application=new OSF.DDA.Application(a);OSF.DDA.DispIdHost.addAsyncMethods(a.application,[OSF.DDA.AsyncMethodNames.CreateDocumentAsync]);OSF.OUtil.finalizeProperties(a.application);OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(a,a.doc,c,null,OSF.DDA.OfficeTheme.getOfficeTheme));OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.DispIdHost.getClientDelegateMethods,OSF.DDA.SafeArray.Delegate.ParameterMap));var b=new OSF.DDA.Theming.InternalThemeHandler;b.InitializeAndChangeOnce()}



var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExtension;
(function (OfficeExtension) {
	var Action=(function () {
		function Action(actionInfo, isWriteOperation, isRestrictedResourceAccess) {
			this.m_actionInfo=actionInfo;
			this.m_isWriteOperation=isWriteOperation;
			this.m_isRestrictedResourceAccess=isRestrictedResourceAccess;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	}());
	OfficeExtension.Action=Action;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TraceMarkerActionResultHandler=(function () {
		function TraceMarkerActionResultHandler(callback) {
			this.m_callback=callback;
		}
		TraceMarkerActionResultHandler.prototype._handleResult=function (value) {
			if (this.m_callback) {
				this.m_callback();
			}
		};
		return TraceMarkerActionResultHandler;
	}());
	var ActionFactory=(function () {
		function ActionFactory() {
		}
		ActionFactory.createSetPropertyAction=function (context, parent, propertyName, value) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args, isRestrictedResourceAccess) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var isWriteOperation=operationType !=1;
			var ret=new OfficeExtension.Action(actionInfo, isWriteOperation, isRestrictedResourceAccess);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createQueryAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createRecursiveQueryAction=function (context, parent, query) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 6,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				RecursiveQueryInfo: query
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createQueryAsJsonAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 7,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createEnsureUnchangedAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 8,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createUpdateAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 9,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			OfficeExtension.Utility.validateObjectPath(obj);
			context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1,
				Name: "",
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			ret.referencedObjectPath=obj._objectPath;
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(obj._objectPath);
			context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
			return ret;
		};
		ActionFactory.createTraceAction=function (context, message, addTraceMessage) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5,
				Name: "Trace",
				ObjectPathId: 0
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			if (addTraceMessage) {
				context._pendingRequest.addTrace(actionInfo.Id, message);
			}
			return ret;
		};
		ActionFactory.createTraceMarkerForCallback=function (context, callback) {
			var action=ActionFactory.createTraceAction(context, null, false);
			context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
		};
		return ActionFactory;
	}());
	OfficeExtension.ActionFactory=ActionFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject=(function () {
		function ClientObject(context, objectPath) {
			OfficeExtension.Utility.checkArgumentNull(context, "context");
			this.m_context=context;
			this.m_objectPath=objectPath;
			if (this.m_objectPath) {
				if (!context._processingResult) {
					OfficeExtension.ActionFactory.createInstantiateAction(context, this);
					if ((context._autoCleanup) && (this._KeepReference)) {
						context.trackedObjects._autoAdd(this);
					}
				}
			}
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			set: function (value) {
				this.m_objectPath=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNullObject", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_isNull", {
			get: function () {
				return this.m_isNull;
			},
			set: function (value) {
				this.m_isNull=value;
				if (value && this.m_objectPath) {
					this.m_objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._handleResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleIdResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			OfficeExtension.Utility.fixObjectPathIfNecessary(this, value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleRetrieveResult=function (value, result) {
			this._handleIdResult(value);
		};
		ClientObject.prototype._recursivelySet=function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
			var isClientObject=(input instanceof ClientObject);
			var originalInput=input;
			if (isClientObject) {
				if (Object.getPrototypeOf(this)===Object.getPrototypeOf(input)) {
					input=JSON.parse(JSON.stringify(input));
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
						argumentName: 'properties',
						errorLocation: this._className+".set"
					});
				}
			}
			try {
				var prop;
				for (var i=0; i < scalarWriteablePropertyNames.length; i++) {
					prop=scalarWriteablePropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							this[prop]=input[prop];
						}
					}
				}
				for (var i=0; i < objectPropertyNames.length; i++) {
					prop=objectPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							var dataToPassToSet=isClientObject ? originalInput[prop] : input[prop];
							this[prop].set(dataToPassToSet, options);
						}
					}
				}
				var throwOnReadOnly=!isClientObject;
				if (options && !OfficeExtension.Utility.isNullOrUndefined(throwOnReadOnly)) {
					throwOnReadOnly=options.throwOnReadOnly;
				}
				for (var i=0; i < notAllowedToBeSetPropertyNames.length; i++) {
					prop=notAllowedToBeSetPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined" && throwOnReadOnly) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
				for (prop in input) {
					if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
						var propertyDescriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
						if (!propertyDescriptor) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
						if (throwOnReadOnly && !propertyDescriptor.set) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".set"
					},
					innerError: innerError
				});
			}
		};
		ClientObject.prototype._recursivelyUpdate=function (properties) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!OfficeExtension.Utility.isSetSupported("RichApiRuntime", "1.2");
			}
			try {
				var scalarPropNames=this[OfficeExtension.Constants.scalarPropertyNames];
				if (!scalarPropNames) {
					scalarPropNames=[];
				}
				var scalarPropUpdatable=this[OfficeExtension.Constants.scalarPropertyUpdateable];
				if (!scalarPropUpdatable) {
					scalarPropUpdatable=[];
					for (var i=0; i < scalarPropNames.length; i++) {
						scalarPropUpdatable.push(false);
					}
				}
				var navigationPropNames=this[OfficeExtension.Constants.navigationPropertyNames];
				if (!navigationPropNames) {
					navigationPropNames=[];
				}
				var scalarProps={};
				var navigationProps={};
				var scalarPropCount=0;
				for (var propName in properties) {
					var index=scalarPropNames.indexOf(propName);
					if (index >=0) {
						if (!scalarPropUpdatable[index]) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, propName),
								debugInfo: {
									errorLocation: propName
								}
							});
						}
						scalarProps[propName]=properties[propName];
++scalarPropCount;
					}
					else if (navigationPropNames.indexOf(propName) >=0) {
						navigationProps[propName]=properties[propName];
					}
					else {
						throw new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.invalidArgument,
							message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, propName),
							debugInfo: {
								errorLocation: propName
							}
						});
					}
				}
				if (scalarPropCount > 0) {
					if (shouldPolyfill) {
						for (var i=0; i < scalarPropNames.length; i++) {
							var propName=scalarPropNames[i];
							var propValue=scalarProps[propName];
							if (!OfficeExtension.Utility.isUndefined(propValue)) {
								OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, propName, propValue);
							}
						}
					}
					else {
						OfficeExtension.ActionFactory.createUpdateAction(this.context, this, scalarProps);
					}
				}
				for (var propName in navigationProps) {
					var navigationPropProxy=this[propName];
					var navigationPropValue=navigationProps[propName];
					navigationPropProxy._recursivelyUpdate(navigationPropValue);
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".update"
					},
					innerError: innerError
				});
			}
		};
		return ClientObject;
	}());
	OfficeExtension.ClientObject=ClientObject;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequest=(function () {
		function ClientRequest(context) {
			this.m_context=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_instantiatedObjectPaths={};
			this.m_flags=0;
			this.m_traceInfos={};
			this.m_pendingProcessEventHandlers=[];
			this.m_pendingEventHandlerActions={};
			this.m_responseTraceIds={};
			this.m_responseTraceMessages=[];
			this.m_preSyncPromises=[];
		}
		Object.defineProperty(ClientRequest.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "traceInfos", {
			get: function () {
				return this.m_traceInfos;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
			get: function () {
				return this.m_responseTraceMessages;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
			get: function () {
				return this.m_responseTraceIds;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._setResponseTraceIds=function (value) {
			if (value) {
				for (var i=0; i < value.length; i++) {
					var traceId=value[i];
					this.m_responseTraceIds[traceId]=traceId;
					var message=this.m_traceInfos[traceId];
					if (!OfficeExtension.Utility.isNullOrUndefined(message)) {
						this.m_responseTraceMessages.push(message);
					}
				}
			}
		};
		ClientRequest.prototype.addAction=function (action) {
			if (this.m_context.batchMode===1) {
				var isSafeAction=false;
				if (action.actionInfo.ActionType===1 &&
					action.referencedObjectPath.objectPathInfo.ObjectPathType===4) {
					isSafeAction=true;
				}
				if (!isSafeAction) {
					this.m_context.ensureInProgressBatchIfBatchMode();
				}
			}
			if (action.isWriteOperation) {
				this.m_flags=this.m_flags | 1;
			}
			if (action.isRestrictedResourceAccess) {
				this.m_flags=this.m_flags | 2;
			}
			this.m_actions.push(action);
			if (action.actionInfo.ActionType==1) {
				this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId]=action;
			}
		};
		Object.defineProperty(ClientRequest.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getLastAction=function () {
			return this.m_actions[this.m_actions.length - 1];
		};
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype.ensureInstantiateObjectPath=function (objectPath) {
			if (objectPath) {
				if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					return;
				}
				this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
				this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
				if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					var actionInfo={
						Id: this.m_context._nextId(),
						ActionType: 1,
						Name: "",
						ObjectPathId: objectPath.objectPathInfo.Id
					};
					var instantiateAction=new OfficeExtension.Action(actionInfo, false, false);
					instantiateAction.referencedObjectPath=objectPath;
					this.addReferencedObjectPath(objectPath);
					this.addAction(instantiateAction);
				}
			}
		};
		ClientRequest.prototype.ensureInstantiateObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.ensureInstantiateObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addReferencedObjectPath=function (objectPath) {
			if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidObjectPath,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath)),
					debugInfo: {
						errorLocation: OfficeExtension.Utility.getObjectPathExpression(objectPath)
					}
				});
			}
			while (objectPath) {
				if (objectPath.isWriteOperation) {
					this.m_flags=this.m_flags | 1;
				}
				if (objectPath.isRestrictedResourceAccess) {
					this.m_flags=this.m_flags | 2;
				}
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequest.prototype.buildRequestMessageBody=function () {
			if (OfficeExtension._internalConfig.enableEarlyDispose) {
				ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			for (var index=0; index < this.m_actions.length; index++) {
				actions.push(this.m_actions[index].actionInfo);
			}
			var ret={
				AutoKeepReference: this.m_context._autoCleanup,
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return ret;
		};
		ClientRequest.prototype.processResponse=function (actionResults) {
			if (actionResults) {
				for (var i=0; i < actionResults.length; i++) {
					var actionResult=actionResults[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequest.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		ClientRequest.prototype._addPendingEventHandlerAction=function (eventHandlers, action) {
			if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
				this.m_pendingEventHandlerActions[eventHandlers._id]=[];
				this.m_pendingProcessEventHandlers.push(eventHandlers);
			}
			this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
			get: function () {
				return this.m_pendingProcessEventHandlers;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getPendingEventHandlerActions=function (eventHandlers) {
			return this.m_pendingEventHandlerActions[eventHandlers._id];
		};
		ClientRequest.prototype._addPreSyncPromise=function (value) {
			this.m_preSyncPromises.push(value);
		};
		Object.defineProperty(ClientRequest.prototype, "_preSyncPromises", {
			get: function () {
				return this.m_preSyncPromises;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_actions", {
			get: function () {
				return this.m_actions;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_objectPaths", {
			get: function () {
				return this.m_referencedObjectPaths;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._removeKeepReferenceAction=function (objectPathId) {
			for (var i=this.m_actions.length - 1; i >=0; i--) {
				var actionInfo=this.m_actions[i].actionInfo;
				if (actionInfo.ObjectPathId===objectPathId && actionInfo.ActionType===3 && actionInfo.Name===OfficeExtension.Constants.keepReference) {
					this.m_actions.splice(i);
					break;
				}
			}
		};
		ClientRequest._updateLastUsedActionIdOfObjectPathId=function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
			while (objectPath) {
				if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
					return;
				}
				lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]=actionId;
				var argumentObjectPaths=objectPath.argumentObjectPaths;
				if (argumentObjectPaths) {
					var argumentObjectPathsLength=argumentObjectPaths.length;
					for (var i=0; i < argumentObjectPathsLength; i++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
					}
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest._calculateLastUsedObjectPathIds=function (actions) {
			var lastUsedActionIdOfObjectPathId={};
			var actionsLength=actions.length;
			for (var index=actionsLength - 1; index >=0; --index) {
				var action=actions[index];
				var actionId=action.actionInfo.Id;
				if (action.referencedObjectPath) {
					ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
				}
				var referencedObjectPaths=action.referencedArgumentObjectPaths;
				if (referencedObjectPaths) {
					var referencedObjectPathsLength=referencedObjectPaths.length;
					for (var refIndex=0; refIndex < referencedObjectPathsLength; refIndex++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
					}
				}
			}
			var lastUsedObjectPathIdsOfAction={};
			for (var key in lastUsedActionIdOfObjectPathId) {
				var actionId=lastUsedActionIdOfObjectPathId[key];
				var objectPathIds=lastUsedObjectPathIdsOfAction[actionId];
				if (!objectPathIds) {
					objectPathIds=[];
					lastUsedObjectPathIdsOfAction[actionId]=objectPathIds;
				}
				objectPathIds.push(parseInt(key));
			}
			for (var index=0; index < actionsLength; index++) {
				var action=actions[index];
				var lastUsedObjectPathIds=lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					action.actionInfo.L=lastUsedObjectPathIds;
				}
				else if (action.actionInfo.L) {
					delete action.actionInfo.L;
				}
			}
		};
		return ClientRequest;
	}());
	OfficeExtension.ClientRequest=ClientRequest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	OfficeExtension._internalConfig={
		showDisposeInfoInDebugInfo: false,
		showInternalApiInDebugInfo: false,
		enableEarlyDispose: true,
		alwaysPolyfillClientObjectUpdateMethod: false,
		alwaysPolyfillClientObjectRetrieveMethod: false
	};
	OfficeExtension.config={
		extendedErrorLogging: false
	};
	var SessionBase=(function () {
		function SessionBase() {
		}
		SessionBase.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		SessionBase.prototype._createRequestExecutorOrNull=function () {
			return null;
		};
		Object.defineProperty(SessionBase.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		return SessionBase;
	}());
	OfficeExtension.SessionBase=SessionBase;
	var ClientRequestContext=(function () {
		function ClientRequestContext(url) {
			this.m_customRequestHeaders={};
			this.m_batchMode=0;
			this._onRunFinishedNotifiers=[];
			this.m_nextId=0;
			if (ClientRequestContext._overrideSession) {
				this.m_requestUrlAndHeaderInfoResolver=ClientRequestContext._overrideSession;
			}
			else {
				if (OfficeExtension.Utility.isNullOrUndefined(url) || typeof (url)==="string" && url.length===0) {
					url=ClientRequestContext.defaultRequestUrlAndHeaders;
					if (!url) {
						url={ url: OfficeExtension.Constants.localDocument, headers: {} };
					}
				}
				if (typeof (url)==="string") {
					this.m_requestUrlAndHeaderInfo={ url: url, headers: {} };
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
					this.m_requestUrlAndHeaderInfoResolver=url;
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
					var requestInfo=url;
					this.m_requestUrlAndHeaderInfo={ url: requestInfo.url, headers: {} };
					OfficeExtension.Utility._copyHeaders(requestInfo.headers, this.m_requestUrlAndHeaderInfo.headers);
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "url" });
				}
			}
			if (this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
				this.m_session=this.m_requestUrlAndHeaderInfoResolver;
			}
			this._processingResult=false;
			this._customData=OfficeExtension.Constants.iterativeExecutor;
			this.sync=this.sync.bind(this);
		}
		Object.defineProperty(ClientRequestContext.prototype, "session", {
			get: function () {
				return this.m_session;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
			get: function () {
				if (this.m_session) {
					return this.m_session.eventRegistration;
				}
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_url", {
			get: function () {
				if (this.m_requestUrlAndHeaderInfo) {
					return this.m_requestUrlAndHeaderInfo.url;
				}
				return null;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
			get: function () {
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
			get: function () {
				var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension._internalConfig.showDisposeInfoInDebugInfo);
				var statements=prettyPrinter.process();
				return { pendingStatements: statements };
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new OfficeExtension.TrackedObjects(this);
				}
				return this.m_trackedObjects;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
			get: function () {
				return this.m_customRequestHeaders;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
			get: function () {
				return this.m_batchMode;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode=function () {
			if (this.m_batchMode===1 && !this.m_explicitBatchInProgress) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.notInsideBatch), null);
			}
		};
		ClientRequestContext.prototype.load=function (clientObj, option) {
			OfficeExtension.Utility.validateContext(this, clientObj);
			var queryOption=ClientRequestContext._parseQueryOption(option);
			var action=OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.isLoadOption=function (loadOption) {
			if (!OfficeExtension.Utility.isUndefined(loadOption.select) && (typeof (loadOption.select)==="string" || Array.isArray(loadOption.select)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.expand) && (typeof (loadOption.expand)==="string" || Array.isArray(loadOption.expand)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.top) && typeof (loadOption.top)==="number")
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.skip) && typeof (loadOption.skip)==="number")
				return true;
			for (var i in loadOption) {
				return false;
			}
			return true;
		};
		ClientRequestContext.parseStrictLoadOption=function (option) {
			var ret={ Select: [] };
			ClientRequestContext.parseStrictLoadOptionHelper(ret, "", "option", option);
			return ret;
		};
		ClientRequestContext.combineQueryPath=function (pathPrefix, key, separator) {
			if (pathPrefix.length===0) {
				return key;
			}
			else {
				return pathPrefix+separator+key;
			}
		};
		ClientRequestContext.parseStrictLoadOptionHelper=function (queryInfo, pathPrefix, argPrefix, option) {
			for (var key in option) {
				var value=option[key];
				if (key==="$all") {
					if (typeof (value) !=="boolean") {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, ".") });
					}
					if (value) {
						queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, "*", "/"));
					}
				}
				else if (key==="$top") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, ".") });
					}
					queryInfo.Top=value;
				}
				else if (key==="$skip") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, ".") });
					}
					queryInfo.Skip=value;
				}
				else {
					if (typeof (value)==="boolean") {
						if (value) {
							queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, key, "/"));
						}
					}
					else if (typeof (value)==="object") {
						ClientRequestContext.parseStrictLoadOptionHelper(queryInfo, ClientRequestContext.combineQueryPath(pathPrefix, key, "/"), ClientRequestContext.combineQueryPath(argPrefix, key, "."), value);
					}
					else {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, ".") });
					}
				}
			}
		};
		ClientRequestContext._parseQueryOption=function (option) {
			var queryOption={};
			if (typeof (option)=="string") {
				var select=option;
				queryOption.Select=OfficeExtension.Utility._parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof (option)==="object") {
				var loadOption=option;
				if (ClientRequestContext.isLoadOption(loadOption)) {
					if (typeof (loadOption.select)=="string") {
						queryOption.Select=OfficeExtension.Utility._parseSelectExpand(loadOption.select);
					}
					else if (Array.isArray(loadOption.select)) {
						queryOption.Select=loadOption.select;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "option.select" });
					}
					if (typeof (loadOption.expand)=="string") {
						queryOption.Expand=OfficeExtension.Utility._parseSelectExpand(loadOption.expand);
					}
					else if (Array.isArray(loadOption.expand)) {
						queryOption.Expand=loadOption.expand;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "option.expand" });
					}
					if (typeof (loadOption.top)==="number") {
						queryOption.Top=loadOption.top;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "option.top" });
					}
					if (typeof (loadOption.skip)==="number") {
						queryOption.Skip=loadOption.skip;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "option.skip" });
					}
				}
				else {
					queryOption=ClientRequestContext.parseStrictLoadOption(option);
				}
			}
			else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "option" });
			}
			return queryOption;
		};
		ClientRequestContext.prototype.loadRecursive=function (clientObj, options, maxDepth) {
			if (!OfficeExtension.Utility.isPlainJsonObject(options)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "options" });
			}
			var quries={};
			for (var key in options) {
				quries[key]=ClientRequestContext._parseQueryOption(options[key]);
			}
			var action=OfficeExtension.ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			OfficeExtension.ActionFactory.createTraceAction(this, message, true);
		};
		ClientRequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) {
		};
		ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo=function () {
			var _this=this;
			return OfficeExtension.Utility._createPromiseFromResult(null)
				.then(function () {
				if (!_this.m_requestUrlAndHeaderInfo) {
					return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo()
						.then(function (value) {
						_this.m_requestUrlAndHeaderInfo=value;
						if (!_this.m_requestUrlAndHeaderInfo) {
							_this.m_requestUrlAndHeaderInfo={ url: OfficeExtension.Constants.localDocument, headers: {} };
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
							_this.m_requestUrlAndHeaderInfo.url=OfficeExtension.Constants.localDocument;
						}
						if (!_this.m_requestUrlAndHeaderInfo.headers) {
							_this.m_requestUrlAndHeaderInfo.headers={};
						}
						if (typeof (_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull)==="function") {
							var executor=_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
							if (executor) {
								_this._requestExecutor=executor;
							}
						}
					});
				}
			});
		};
		ClientRequestContext.prototype.syncPrivateMain=function () {
			var _this=this;
			return this.ensureRequestUrlAndHeaderInfo()
				.then(function () {
				var req=_this._pendingRequest;
				_this.m_pendingRequest=null;
				return _this.processPreSyncPromises(req)
					.then(function () { return _this.syncPrivate(req); });
			});
		};
		ClientRequestContext.prototype.syncPrivate=function (req) {
			var _this=this;
			if (!req.hasActions) {
				return this.processPendingEventHandlers(req);
			}
			var msgBody=req.buildRequestMessageBody();
			var requestFlags=req.flags;
			if (!this._requestExecutor) {
				if (OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
					this._requestExecutor=new OfficeExtension.OfficeJsRequestExecutor(this);
				}
				else {
					this._requestExecutor=new OfficeExtension.HttpRequestExecutor();
				}
			}
			var requestExecutor=this._requestExecutor;
			var headers={};
			OfficeExtension.Utility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
			OfficeExtension.Utility._copyHeaders(this.m_customRequestHeaders, headers);
			var requestExecutorRequestMessage={
				Url: this.m_requestUrlAndHeaderInfo.url,
				Headers: headers,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var errorFromResponse=null;
			var errorFromProcessEventHandlers=null;
			this._lastSyncStart=performance.now();
			return requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
				.then(function (response) {
				_this._lastSyncEnd=performance.now();
				errorFromResponse=_this.processRequestExecutorResponseMessage(req, response);
				return _this.processPendingEventHandlers(req)
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error in processPendingEventHandlers");
					OfficeExtension.Utility.log(JSON.stringify(ex));
					errorFromProcessEventHandlers=ex;
				});
			})
				.then(function () {
				if (errorFromResponse) {
					OfficeExtension.Utility.log("Throw error from response: "+JSON.stringify(errorFromResponse));
					throw errorFromResponse;
				}
				if (errorFromProcessEventHandlers) {
					OfficeExtension.Utility.log("Throw error from ProcessEventHandler: "+JSON.stringify(errorFromProcessEventHandlers));
					var transformedError=null;
					if (errorFromProcessEventHandlers instanceof OfficeExtension._Internal.RuntimeError) {
						transformedError=errorFromProcessEventHandlers;
						transformedError.traceMessages=req._responseTraceMessages;
					}
					else {
						var message=null;
						if (typeof (errorFromProcessEventHandlers)==="string") {
							message=errorFromProcessEventHandlers;
						}
						else {
							message=errorFromProcessEventHandlers.message;
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(message)) {
							message=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotRegisterEvent);
						}
						transformedError=new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.cannotRegisterEvent,
							message: message,
							traceMessages: req._responseTraceMessages
						});
					}
					throw transformedError;
				}
			});
		};
		ClientRequestContext.prototype.processRequestExecutorResponseMessage=function (req, response) {
			if (response.Body && response.Body.TraceIds) {
				req._setResponseTraceIds(response.Body.TraceIds);
			}
			var traceMessages=req._responseTraceMessages;
			var errorStatementInfo=null;
			if (response.Body) {
				if (response.Body.Error &&
					response.Body.Error.ActionIndex >=0) {
					var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, true);
					var debugInfoStatementInfo=prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
					errorStatementInfo={
						statement: debugInfoStatementInfo.statement,
						surroundingStatements: debugInfoStatementInfo.surroundingStatements,
						fullStatements: ["Please enable config.extendedErrorLogging to see full statements."]
					};
					if (OfficeExtension.config.extendedErrorLogging) {
						prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, false);
						errorStatementInfo.fullStatements=prettyPrinter.process();
					}
				}
				var actionResults=null;
				if (response.Body.Results) {
					actionResults=response.Body.Results;
				}
				else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
					actionResults=response.Body.ProcessedResults.Results;
				}
				if (actionResults) {
					this._processingResult=true;
					try {
						req.processResponse(actionResults);
					}
					finally {
						this._processingResult=false;
					}
				}
			}
			if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
				return new OfficeExtension._Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage,
					traceMessages: traceMessages
				});
			}
			else if (response.Body && response.Body.Error) {
				var debugInfo={
					errorLocation: response.Body.Error.Location
				};
				if (errorStatementInfo) {
					debugInfo.statement=errorStatementInfo.statement;
					debugInfo.surroundingStatements=errorStatementInfo.surroundingStatements;
					debugInfo.fullStatements=errorStatementInfo.fullStatements;
				}
				return new OfficeExtension._Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message,
					traceMessages: traceMessages,
					debugInfo: debugInfo
				});
			}
			return null;
		};
		ClientRequestContext.prototype.processPendingEventHandlers=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._pendingProcessEventHandlers.length; i++) {
				var eventHandlers=req._pendingProcessEventHandlers[i];
				ret=ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneEventHandlersFunc=function (eventHandlers, req) {
			return function () { return eventHandlers._processRegistration(req); };
		};
		ClientRequestContext.prototype.processPreSyncPromises=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._preSyncPromises.length; i++) {
				var p=req._preSyncPromises[i];
				ret=ret.then(this.createProcessOneProSyncFunc(p));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneProSyncFunc=function (p) {
			return function () { return p; };
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			return this.syncPrivateMain().then(function () { return passThroughValue; });
		};
		ClientRequestContext.prototype.batch=function (batchBody) {
			var _this=this;
			if (this.m_batchMode !==1) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, null, null));
			}
			if (this.m_explicitBatchInProgress) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.pendingBatchInProgress), null));
			}
			if (OfficeExtension.Utility.isNullOrUndefined(batchBody)) {
				return OfficeExtension.Utility._createPromiseFromResult(null);
			}
			this.m_explicitBatchInProgress=true;
			var previousRequest=this.m_pendingRequest;
			this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
			var batchBodyResult;
			try {
				batchBodyResult=batchBody(this._rootObject, this);
			}
			catch (ex) {
				this.m_explicitBatchInProgress=false;
				this.m_pendingRequest=previousRequest;
				return OfficeExtension._Internal.OfficePromise.reject(ex);
			}
			var request;
			var batchBodyResultPromise;
			if (typeof (batchBodyResult)==="object" &&
				batchBodyResult &&
				typeof (batchBodyResult.then)==="function") {
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(function () {
					return batchBodyResult;
				})
					.then(function (result) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return result;
				})
					.catch(function (ex) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return OfficeExtension._Internal.OfficePromise.reject(ex);
				});
			}
			else {
				this.m_explicitBatchInProgress=false;
				request=this.m_pendingRequest;
				this.m_pendingRequest=previousRequest;
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(batchBodyResult);
			}
			return batchBodyResultPromise
				.then(function (result) {
				return _this.ensureRequestUrlAndHeaderInfo()
					.then(function () {
					return _this.syncPrivate(request);
				})
					.then(function () {
					return result;
				});
			});
		};
		ClientRequestContext._run=function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runCommon("run", null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.isRequestUrlAndHeaderInfo=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				Object.getPrototypeOf(value)===Object.getPrototypeOf({}) &&
				!OfficeExtension.Utility.isNullOrUndefined(value.url));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfoResolver=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				typeof (value._resolveRequestUrlAndHeaderInfo)==="function");
		};
		ClientRequestContext._runBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runExplicitBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runBatchCommon=function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			var ctxRetriever;
			var batch;
			var requestInfo=null;
			var argOffset=0;
			if (receivedRunArgs.length > 0 &&
				(typeof (receivedRunArgs[0])==="string" ||
					ClientRequestContext.isRequestUrlAndHeaderInfo(receivedRunArgs[0]) ||
					ClientRequestContext.isRequestUrlAndHeaderInfoResolver(receivedRunArgs[0]))) {
				requestInfo=receivedRunArgs[0];
				argOffset=1;
			}
			if (receivedRunArgs.length==argOffset+1) {
				ctxRetriever=ctxInitializer;
				batch=receivedRunArgs[argOffset+0];
			}
			else if (receivedRunArgs.length==argOffset+2) {
				if (OfficeExtension.Utility.isNullOrUndefined(receivedRunArgs[argOffset+0])) {
					ctxRetriever=ctxInitializer;
				}
				else if (receivedRunArgs[argOffset+0] instanceof OfficeExtension.ClientObject) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0].context; };
				}
				else if (receivedRunArgs[argOffset+0] instanceof ClientRequestContext) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0]; };
				}
				else if (Array.isArray(receivedRunArgs[argOffset+0])) {
					var array=receivedRunArgs[argOffset+0];
					if (array.length==0) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					for (var i=0; i < array.length; i++) {
						if (!(array[i] instanceof OfficeExtension.ClientObject)) {
							return ClientRequestContext.createErrorPromise(functionName);
						}
						if (array[i].context !=array[0].context) {
							return ClientRequestContext.createErrorPromise(functionName, OfficeExtension.ResourceStrings.invalidRequestContext);
						}
					}
					ctxRetriever=function () { return array[0].context; };
				}
				else {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				batch=receivedRunArgs[argOffset+1];
			}
			else {
				return ClientRequestContext.createErrorPromise(functionName);
			}
			return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.createErrorPromise=function (functionName, code) {
			if (code===void 0) { code=OfficeExtension.ResourceStrings.invalidArgument; }
			return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(code, OfficeExtension.Utility._getResourceString(code), functionName));
		};
		ClientRequestContext._runCommon=function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (ClientRequestContext._overrideSession) {
				requestInfo=ClientRequestContext._overrideSession;
			}
			var starterPromise=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) { resolve(); });
			var ctx;
			var succeeded=false;
			var resultOrError;
			var previousBatchMode;
			return starterPromise
				.then(function () {
				ctx=ctxRetriever(requestInfo);
				if (ctx._autoCleanup) {
					return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
						ctx._onRunFinishedNotifiers.push(function () {
							ctx._autoCleanup=true;
							resolve();
						});
					});
				}
				else {
					ctx._autoCleanup=true;
				}
			})
				.then(function () {
				if (typeof runBody !=='function') {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				previousBatchMode=ctx.m_batchMode;
				ctx.m_batchMode=batchMode;
				var runBodyResult;
				if (batchMode==1) {
					runBodyResult=runBody(ctx.batch.bind(ctx));
				}
				else {
					runBodyResult=runBody(ctx);
				}
				if (OfficeExtension.Utility.isNullOrUndefined(runBodyResult) || (typeof runBodyResult.then !=='function')) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
				}
				return runBodyResult;
			})
				.then(function (runBodyResult) {
				if (batchMode===1) {
					return runBodyResult;
				}
				else {
					return ctx.sync(runBodyResult);
				}
			})
				.then(function (result) {
				succeeded=true;
				resultOrError=result;
			})
				.catch(function (error) {
				resultOrError=error;
			})
				.then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				ctx.m_batchMode=previousBatchMode;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				if (OfficeExtension.Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
					return attemptCleanup();
				}
				else {
					attemptCleanup();
				}
				function attemptCleanup() {
					cleanupCounter++;
					var savedPendingRequest=ctx.m_pendingRequest;
					var savedBatchMode=ctx.m_batchMode;
					var request=new OfficeExtension.ClientRequest(ctx);
					ctx.m_pendingRequest=request;
					ctx.m_batchMode=0;
					try {
						for (var key in itemsToRemove) {
							ctx.trackedObjects.remove(itemsToRemove[key]);
						}
					}
					finally {
						ctx.m_batchMode=savedBatchMode;
						ctx.m_pendingRequest=savedPendingRequest;
					}
					return ctx.syncPrivate(request)
						.then(function () {
						if (onCleanupSuccess) {
							onCleanupSuccess(cleanupCounter);
						}
					})
						.catch(function () {
						if (onCleanupFailure) {
							onCleanupFailure(cleanupCounter);
						}
						if (cleanupCounter < numCleanupAttempts) {
							setTimeout(function () {
								attemptCleanup();
							}, retryDelay);
						}
					});
				}
			})
				.then(function () {
				if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
					var func=ctx._onRunFinishedNotifiers.shift();
					func();
				}
				if (succeeded) {
					return resultOrError;
				}
				else {
					throw resultOrError;
				}
			});
		};
		ClientRequestContext.prototype._nextId=function () {
			return++this.m_nextId;
		};
		return ClientRequestContext;
	}());
	OfficeExtension.ClientRequestContext=ClientRequestContext;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult=(function () {
		function ClientResult(type) {
			this.m_type=type;
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "clientResult.value"
						}
					});
				}
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (typeof (value)==="object" && value && value._IsNull) {
				return;
			}
			if (this.m_type===1) {
				this.m_value=OfficeExtension.Utility.adjustToDateTime(value);
			}
			else {
				this.m_value=value;
			}
		};
		return ClientResult;
	}());
	OfficeExtension.ClientResult=ClientResult;
	var RetrieveResultImpl=(function () {
		function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
			this.m_proxy=m_proxy;
			this.m_shouldPolyfill=m_shouldPolyfill;
			var scalarPropertyNames=m_proxy[OfficeExtension.Constants.scalarPropertyNames];
			var navigationPropertyNames=m_proxy[OfficeExtension.Constants.navigationPropertyNames];
			var typeName=m_proxy[OfficeExtension.Constants.className];
			var isCollection=m_proxy[OfficeExtension.Constants.isCollection];
			if (scalarPropertyNames) {
				for (var i=0; i < scalarPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
				}
			}
			if (navigationPropertyNames) {
				for (var i=0; i < navigationPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
				}
			}
			if (isCollection) {
				OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, OfficeExtension.Constants.itemsLowerCase);
			}
		}
		Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
			get: function () {
				return this.m_proxy;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "retrieveResult.$isNullObject"
						}
					});
				}
				return this.m_isNullObject;
			},
			enumerable: true,
			configurable: true
		});
		RetrieveResultImpl.prototype.toJSON=function () {
			if (!this.m_isLoaded) {
				return undefined;
			}
			if (this.m_isNullObject) {
				return null;
			}
			if (OfficeExtension.Utility.isUndefined(this.m_json)) {
				this.m_json=this.purifyJson(this.m_value);
			}
			return this.m_json;
		};
		RetrieveResultImpl.prototype.toString=function () {
			return JSON.stringify(this.toJSON());
		};
		RetrieveResultImpl.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (value===null || typeof (value)==="object" && value && value._IsNull) {
				this.m_isNullObject=true;
				value=null;
			}
			else {
				this.m_isNullObject=false;
			}
			if (this.m_shouldPolyfill) {
				value=this.changePropertyNameToCamelLowerCase(value);
			}
			this.m_value=value;
			this.m_proxy._handleRetrieveResult(value, this);
		};
		RetrieveResultImpl.prototype.changePropertyNameToCamelLowerCase=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					var propValue=value[key];
					if (key===OfficeExtension.Constants.items) {
						ret={};
						ret[OfficeExtension.Constants.itemsLowerCase]=this.changePropertyNameToCamelLowerCase(propValue);
						break;
					}
					else {
						var propName=OfficeExtension.Utility._toCamelLowerCase(key);
						ret[propName]=this.changePropertyNameToCamelLowerCase(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		RetrieveResultImpl.prototype.purifyJson=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.purifyJson(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					if (key.charCodeAt(0) !==charCodeUnderscore) {
						var propValue=value[key];
						if (typeof (propValue)==="object" &&
							propValue !==null &&
							Array.isArray(propValue["items"])) {
							propValue=propValue["items"];
						}
						ret[key]=this.purifyJson(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		return RetrieveResultImpl;
	}());
	OfficeExtension.RetrieveResultImpl=RetrieveResultImpl;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Constants=(function () {
		function Constants() {
		}
		Constants.flags="flags";
		Constants.getItemAt="GetItemAt";
		Constants.id="Id";
		Constants.idLowerCase="id";
		Constants.idPrivate="_Id";
		Constants.index="_Index";
		Constants.items="_Items";
		Constants.iterativeExecutor="IterativeExecutor";
		Constants.localDocument="http://document.localhost/";
		Constants.localDocumentApiPrefix="http://document.localhost/_api/";
		Constants.keepReference="_KeepReference";
		Constants.objectPathIdPrivate="_ObjectPathId";
		Constants.processQuery="ProcessQuery";
		Constants.referenceId="_ReferenceId";
		Constants.isTracked="_IsTracked";
		Constants.sourceLibHeader="SdkVersion";
		Constants.sessionContext="sc";
		Constants.embeddingPageOrigin="EmbeddingPageOrigin";
		Constants.embeddingPageSessionInfo="EmbeddingPageSessionInfo";
		Constants.eventMessageCategory=65536;
		Constants.eventWorkbookId="Workbook";
		Constants.eventSourceRemote="Remote";
		Constants.itemsLowerCase="items";
		Constants.proxy="$proxy";
		Constants.scalarPropertyNames="_scalarPropertyNames";
		Constants.navigationPropertyNames="_navigationPropertyNames";
		Constants.className="_className";
		Constants.isCollection="_isCollection";
		Constants.scalarPropertyUpdateable="_scalarPropertyUpdateable";
		Constants.collectionPropertyPath="_collectionPropertyPath";
		Constants.objectPathInfoDoNotKeepReferenceFieldName="D";
		return Constants;
	}());
	OfficeExtension.Constants=Constants;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var versionToken=1;
	var internalConfiguration={
		invokeRequestModifier: function (request) {
			request.DdaMethod.Version=versionToken;
			return request;
		},
		invokeResponseModifier: function (args) {
			versionToken=args.Version;
			if (args.Error) {
				args.error={};
				args.error.Code=args.Error;
			}
			return args;
		}
	};
	var EmbeddedApiStatus;
	(function (EmbeddedApiStatus) {
		EmbeddedApiStatus[EmbeddedApiStatus["Success"]=0]="Success";
		EmbeddedApiStatus[EmbeddedApiStatus["Timeout"]=1]="Timeout";
		EmbeddedApiStatus[EmbeddedApiStatus["InternalError"]=5001]="InternalError";
	})(EmbeddedApiStatus || (EmbeddedApiStatus={}));
	var CommunicationConstants;
	(function (CommunicationConstants) {
		CommunicationConstants.SendingId="sId";
		CommunicationConstants.RespondingId="rId";
		CommunicationConstants.CommandKey="command";
		CommunicationConstants.SessionInfoKey="sessionInfo";
		CommunicationConstants.ParamsKey="params";
		CommunicationConstants.ApiReadyCommand="apiready";
		CommunicationConstants.ExecuteMethodCommand="executeMethod";
		CommunicationConstants.GetAppContextCommand="getAppContext";
		CommunicationConstants.RegisterEventCommand="registerEvent";
		CommunicationConstants.UnregisterEventCommand="unregisterEvent";
		CommunicationConstants.FireEventCommand="fireEvent";
	})(CommunicationConstants || (CommunicationConstants={}));
	var EmbeddedSession=(function (_super) {
		__extends(EmbeddedSession, _super);
		function EmbeddedSession(url, options) {
			_super.call(this);
			this.m_chosenWindow=null;
			this.m_chosenOrigin=null;
			this.m_enabled=true;
			this.m_onMessageHandler=this._onMessage.bind(this);
			this.m_callbackList={};
			this.m_id=0;
			this.m_timeoutId=-1;
			this.m_appContext=null;
			this.m_url=url;
			this.m_options=options;
			if (!this.m_options) {
				this.m_options={ sessionKey: Math.random().toString() };
			}
			if (!this.m_options.sessionKey) {
				this.m_options.sessionKey=Math.random().toString();
			}
			if (!this.m_options.container) {
				this.m_options.container=document.body;
			}
			if (!this.m_options.timeoutInMilliseconds) {
				this.m_options.timeoutInMilliseconds=60000;
			}
			if (!this.m_options.height) {
				this.m_options.height="400px";
			}
			if (!this.m_options.width) {
				this.m_options.width="100%";
			}
			if (!(this.m_options.webApplication && this.m_options.webApplication.accessToken && this.m_options.webApplication.accessTokenTtl)) {
				this.m_options.webApplication=null;
			}
		}
		EmbeddedSession.prototype._getIFrameSrc=function () {
			var origin=window.location.protocol+"//"+window.location.host;
			var toAppend=OfficeExtension.Constants.embeddingPageOrigin+"="+encodeURIComponent(origin)+"&"+OfficeExtension.Constants.embeddingPageSessionInfo+"="+encodeURIComponent(this.m_options.sessionKey);
			var useHash=false;
			if (this.m_url.toLowerCase().indexOf("/_layouts/preauth.aspx") > 0 ||
				this.m_url.toLowerCase().indexOf("/_layouts/15/preauth.aspx") > 0) {
				useHash=true;
			}
			var a=document.createElement("a");
			a.href=this.m_url;
			if (this.m_options.webApplication) {
				var toAppendWAC=OfficeExtension.Constants.embeddingPageOrigin+"="+origin+"&"+OfficeExtension.Constants.embeddingPageSessionInfo+"="+this.m_options.sessionKey;
				if (a.search.length===0 || a.search==="?") {
					a.search="?"+OfficeExtension.Constants.sessionContext+"="+encodeURIComponent(toAppendWAC);
				}
				else {
					a.search=a.search+"&"+OfficeExtension.Constants.sessionContext+"="+encodeURIComponent(toAppendWAC);
				}
			}
			else if (useHash) {
				if (a.hash.length===0 || a.hash==="#") {
					a.hash="#"+toAppend;
				}
				else {
					a.hash=a.hash+"&"+toAppend;
				}
			}
			else {
				if (a.search.length===0 || a.search==="?") {
					a.search="?"+toAppend;
				}
				else {
					a.search=a.search+"&"+toAppend;
				}
			}
			var iframeSrc=a.href;
			return iframeSrc;
		};
		EmbeddedSession.prototype.init=function () {
			var _this=this;
			window.addEventListener("message", this.m_onMessageHandler);
			var iframeSrc=this._getIFrameSrc();
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var iframeElement=document.createElement("iframe");
				if (_this.m_options.id) {
					iframeElement.id=_this.m_options.id;
					iframeElement.name=_this.m_options.id;
				}
				iframeElement.style.height=_this.m_options.height;
				iframeElement.style.width=_this.m_options.width;
				if (!_this.m_options.webApplication) {
					iframeElement.src=iframeSrc;
					_this.m_options.container.appendChild(iframeElement);
				}
				else {
					var webApplicationForm=document.createElement('form');
					webApplicationForm.setAttribute("action", iframeSrc);
					webApplicationForm.setAttribute("method", "post");
					webApplicationForm.setAttribute("target", iframeElement.name);
					_this.m_options.container.appendChild(webApplicationForm);
					var token_input=document.createElement('input');
					token_input.setAttribute("type", "hidden");
					token_input.setAttribute("name", "access_token");
					token_input.setAttribute("value", _this.m_options.webApplication.accessToken);
					webApplicationForm.appendChild(token_input);
					var token_ttl_input=document.createElement('input');
					token_ttl_input.setAttribute("type", "hidden");
					token_ttl_input.setAttribute("name", "access_token_ttl");
					token_ttl_input.setAttribute("value", _this.m_options.webApplication.accessTokenTtl);
					webApplicationForm.appendChild(token_ttl_input);
					_this.m_options.container.appendChild(iframeElement);
					webApplicationForm.submit();
				}
				_this.m_timeoutId=setTimeout(function () {
					_this.close();
					var err=OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.timeout, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.timeout), "EmbeddedSession.init");
					reject(err);
				}, _this.m_options.timeoutInMilliseconds);
				_this.m_promiseResolver=resolve;
			});
		};
		EmbeddedSession.prototype._invoke=function (method, callback, params) {
			if (!this.m_enabled) {
				callback(EmbeddedApiStatus.InternalError, null);
				return;
			}
			if (internalConfiguration.invokeRequestModifier) {
				params=internalConfiguration.invokeRequestModifier(params);
			}
			this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
				if (internalConfiguration.invokeResponseModifier) {
					args=internalConfiguration.invokeResponseModifier(args);
				}
				var errorCode=args["Error"];
				delete args["Error"];
				callback(errorCode || EmbeddedApiStatus.Success, args);
			});
		};
		EmbeddedSession.prototype.close=function () {
			window.removeEventListener("message", this.m_onMessageHandler);
			window.clearTimeout(this.m_timeoutId);
			this.m_enabled=false;
		};
		Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
			get: function () {
				if (!this.m_sessionEventManager) {
					this.m_sessionEventManager=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
				}
				return this.m_sessionEventManager;
			},
			enumerable: true,
			configurable: true
		});
		EmbeddedSession.prototype._createRequestExecutorOrNull=function () {
			return new EmbeddedRequestExecutor(this);
		};
		EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EmbeddedSession.prototype._registerEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve(null);
				});
			});
		};
		EmbeddedSession.prototype._unregisterEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve();
				});
			});
		};
		EmbeddedSession.prototype._onMessage=function (event) {
			var _this=this;
			if (!this.m_enabled) {
				return;
			}
			if (this.m_chosenWindow
				&& (this.m_chosenWindow !==event.source || this.m_chosenOrigin !==event.origin)) {
				return;
			}
			var eventData=event.data;
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.ApiReadyCommand) {
				if (!this.m_chosenWindow
					&& this._isValidDescendant(event.source)
					&& eventData[CommunicationConstants.SessionInfoKey]===this.m_options.sessionKey) {
					this.m_chosenWindow=event.source;
					this.m_chosenOrigin=event.origin;
					this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
						_this._setupContext(appContext);
						window.clearTimeout(_this.m_timeoutId);
						_this.m_promiseResolver();
					});
				}
				return;
			}
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.FireEventCommand) {
				var msg=eventData[CommunicationConstants.ParamsKey];
				var eventId=msg["EventId"];
				var targetId=msg["TargetId"];
				var data=msg["Data"];
				if (this.m_sessionEventManager) {
					var handlers=this.m_sessionEventManager.getHandlers(eventId, targetId);
					for (var i=0; i < handlers.length; i++) {
						handlers[i](data);
					}
				}
				return;
			}
			if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
				var rId=eventData[CommunicationConstants.RespondingId];
				var callback=this.m_callbackList[rId];
				if (typeof callback==="function") {
					callback(eventData[CommunicationConstants.ParamsKey]);
				}
				delete this.m_callbackList[rId];
			}
		};
		EmbeddedSession.prototype._sendMessageWithCallback=function (id, command, data, callback) {
			this.m_callbackList[id]=callback;
			var message={};
			message[CommunicationConstants.SendingId]=id;
			message[CommunicationConstants.CommandKey]=command;
			message[CommunicationConstants.ParamsKey]=data;
			this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
		};
		EmbeddedSession.prototype._isValidDescendant=function (wnd) {
			var container=this.m_options.container || document.body;
			function doesFrameWindow(containerWindow) {
				if (containerWindow===wnd) {
					return true;
				}
				for (var i=0, len=containerWindow.frames.length; i < len; i++) {
					if (doesFrameWindow(containerWindow.frames[i])) {
						return true;
					}
				}
				return false;
			}
			var iframes=container.getElementsByTagName("iframe");
			for (var i=0, len=iframes.length; i < len; i++) {
				if (doesFrameWindow(iframes[i].contentWindow)) {
					return true;
				}
			}
			return false;
		};
		EmbeddedSession.prototype._setupContext=function (appContext) {
			if (!(this.m_appContext=appContext)) {
				return;
			}
		};
		return EmbeddedSession;
	}(OfficeExtension.SessionBase));
	OfficeExtension.EmbeddedSession=EmbeddedSession;
	var EmbeddedRequestExecutor=(function () {
		function EmbeddedRequestExecutor(session) {
			this.m_session=session;
		}
		EmbeddedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (status==EmbeddedApiStatus.Success) {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
					}
					resolve(response);
				}, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
			});
		};
		EmbeddedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbeddedRequestExecutor.SourceLibHeaderValue="Embedded";
		return EmbeddedRequestExecutor;
	}());
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(error) {
				_super.call(this, (typeof error==="string") ? error : error.message);
				this.name="OfficeExtension.Error";
				if (typeof error==="string") {
					this.message=error;
				}
				else {
					this.code=error.code;
					this.message=error.message;
					this.traceMessages=error.traceMessages || [];
					this.innerError=error.innerError || null;
					this.debugInfo=this._createDebugInfo(error.debugInfo || {});
				}
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			RuntimeError.prototype._createDebugInfo=function (partialDebugInfo) {
				var debugInfo={
					code: this.code,
					message: this.message,
					toString: function () {
						return JSON.stringify(this);
					}
				};
				for (var key in partialDebugInfo) {
					debugInfo[key]=partialDebugInfo[key];
				}
				if (this.innerError) {
					if (this.innerError instanceof OfficeExtension._Internal.RuntimeError) {
						debugInfo.innerError=this.innerError.debugInfo;
					}
					else {
						debugInfo.innerError=this.innerError;
					}
				}
				return debugInfo;
			};
			RuntimeError._createInvalidArgError=function (error) {
				return new _Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: (OfficeExtension.Utility.isNullOrEmptyString(error.argumentName) ?
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgumentGeneric) :
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, error.argumentName)),
					debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
					innerError: error.innerError
				});
			};
			return RuntimeError;
		}(Error));
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	OfficeExtension.Error=_Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes=(function () {
		function ErrorCodes() {
		}
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.activityLimitReached="ActivityLimitReached";
		ErrorCodes.invalidObjectPath="InvalidObjectPath";
		ErrorCodes.propertyNotLoaded="PropertyNotLoaded";
		ErrorCodes.valueNotLoaded="ValueNotLoaded";
		ErrorCodes.invalidRequestContext="InvalidRequestContext";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.runMustReturnPromise="RunMustReturnPromise";
		ErrorCodes.cannotRegisterEvent="CannotRegisterEvent";
		ErrorCodes.apiNotFound="ApiNotFound";
		ErrorCodes.connectionFailure="ConnectionFailure";
		ErrorCodes.timeout="Timeout";
		ErrorCodes.invalidOrTimedOutSession="InvalidOrTimedOutSession";
		ErrorCodes.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ErrorCodes;
	}());
	OfficeExtension.ErrorCodes=ErrorCodes;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlers=(function () {
		function EventHandlers(context, parentObject, name, eventInfo) {
			var _this=this;
			this.m_id=context._nextId();
			this.m_context=context;
			this.m_name=name;
			this.m_handlers=[];
			this.m_registered=false;
			this.m_eventInfo=eventInfo;
			this.m_callback=function (args) {
				_this.m_eventInfo.eventArgsTransformFunc(args)
					.then(function (newArgs) { return _this.fireEvent(newArgs); });
			};
		}
		Object.defineProperty(EventHandlers.prototype, "_registered", {
			get: function () {
				return this.m_registered;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_id", {
			get: function () {
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_handlers", {
			get: function () {
				return this.m_handlers;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_callback", {
			get: function () {
				return this.m_callback;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlers.prototype.add=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 0 });
			return new OfficeExtension.EventHandlerResult(this.m_context, this, handler);
		};
		EventHandlers.prototype.remove=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 1 });
		};
		EventHandlers.prototype.removeAll=function () {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: null, operation: 2 });
		};
		EventHandlers.prototype._processRegistration=function (req) {
			var _this=this;
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			var actions=req._getPendingEventHandlerActions(this);
			if (!actions) {
				return ret;
			}
			var handlersResult=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				handlersResult.push(this.m_handlers[i]);
			}
			var hasChange=false;
			for (var i=0; i < actions.length; i++) {
				if (req._responseTraceIds[actions[i].id]) {
					hasChange=true;
					switch (actions[i].operation) {
						case 0:
							handlersResult.push(actions[i].handler);
							break;
						case 1:
							for (var index=handlersResult.length - 1; index >=0; index--) {
								if (handlersResult[index]===actions[i].handler) {
									handlersResult.splice(index, 1);
									break;
								}
							}
							break;
						case 2:
							handlersResult=[];
							break;
					}
				}
			}
			if (hasChange) {
				if (!this.m_registered && handlersResult.length > 0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); })
						.then(function () { return (_this.m_registered=true); });
				}
				else if (this.m_registered && handlersResult.length==0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
						.catch(function (ex) {
						OfficeExtension.Utility.log("Error when unregister event: "+JSON.stringify(ex));
					})
						.then(function () { return (_this.m_registered=false); });
				}
				ret=ret
					.then(function () { return (_this.m_handlers=handlersResult); });
			}
			return ret;
		};
		EventHandlers.prototype.fireEvent=function (args) {
			var promises=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				var handler=this.m_handlers[i];
				var p=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(this.createFireOneEventHandlerFunc(handler, args))
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error when invoke handler: "+JSON.stringify(ex));
				});
				promises.push(p);
			}
			OfficeExtension._Internal.OfficePromise.all(promises);
		};
		EventHandlers.prototype.createFireOneEventHandlerFunc=function (handler, args) {
			return function () { return handler(args); };
		};
		return EventHandlers;
	}());
	OfficeExtension.EventHandlers=EventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlerResult=(function () {
		function EventHandlerResult(context, handlers, handler) {
			this.m_context=context;
			this.m_allHandlers=handlers;
			this.m_handler=handler;
		}
		Object.defineProperty(EventHandlerResult.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlerResult.prototype.remove=function () {
			if (this.m_allHandlers && this.m_handler) {
				this.m_allHandlers.remove(this.m_handler);
				this.m_allHandlers=null;
				this.m_handler=null;
			}
		};
		return EventHandlerResult;
	}());
	OfficeExtension.EventHandlerResult=EventHandlerResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var OfficeJsEventRegistration=(function () {
			function OfficeJsEventRegistration() {
			}
			OfficeJsEventRegistration.prototype.register=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: "eventId" });
				}
			};
			OfficeJsEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", { handler: handler }, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: "eventId" });
				}
			};
			return OfficeJsEventRegistration;
		}());
		_Internal.officeJsEventRegistration=new OfficeJsEventRegistration();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var EventRegistration=(function () {
		function EventRegistration(registerEventImpl, unregisterEventImpl) {
			this.m_handlersByEventByTarget={};
			this.m_registerEventImpl=registerEventImpl;
			this.m_unregisterEventImpl=unregisterEventImpl;
		}
		EventRegistration.prototype.getHandlers=function (eventId, targetId) {
			if (OfficeExtension.Utility.isNullOrUndefined(targetId)) {
				targetId="";
			}
			var handlersById=this.m_handlersByEventByTarget[eventId];
			if (!handlersById) {
				handlersById={};
				this.m_handlersByEventByTarget[eventId]=handlersById;
			}
			var handlers=handlersById[targetId];
			if (!handlers) {
				handlers=[];
				handlersById[targetId]=handlers;
			}
			return handlers;
		};
		EventRegistration.prototype.register=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: "handler" });
			}
			var handlers=this.getHandlers(eventId, targetId);
			handlers.push(handler);
			if (handlers.length===1) {
				return this.m_registerEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: "handler" });
			}
			var handlers=this.getHandlers(eventId, targetId);
			for (var index=handlers.length - 1; index >=0; index--) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					break;
				}
			}
			if (handlers.length===0) {
				return this.m_unregisterEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		return EventRegistration;
	}());
	OfficeExtension.EventRegistration=EventRegistration;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var GenericEventRegistration=(function () {
		function GenericEventRegistration() {
			this.m_eventRegistration=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
			this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this);
		}
		GenericEventRegistration.prototype.ready=function () {
			var _this=this;
			if (!this.m_ready) {
				if (GenericEventRegistration._testReadyImpl) {
					this.m_ready=GenericEventRegistration._testReadyImpl()
						.then(function () {
						_this.m_isReady=true;
					});
				}
				else {
					this.m_ready=OfficeExtension._Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler)
						.then(function () {
						_this.m_isReady=true;
					});
				}
			}
			return this.m_ready;
		};
		Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
			get: function () {
				return this.m_isReady;
			},
			enumerable: true,
			configurable: true
		});
		GenericEventRegistration.prototype.register=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype._registerEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._unregisterEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._handleRichApiMessage=function (msg) {
			if (msg && msg.entries) {
				for (var entryIndex=0; entryIndex < msg.entries.length; entryIndex++) {
					var entry=msg.entries[entryIndex];
					if (entry.messageCategory==OfficeExtension.Constants.eventMessageCategory) {
						if (OfficeExtension.Utility._logEnabled) {
							OfficeExtension.Utility.log(JSON.stringify(entry));
						}
						var funcs=this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
						if (funcs.length > 0) {
							var arg=JSON.parse(entry.message);
							if (entry.isRemoteOverride) {
								arg.source=OfficeExtension.Constants.eventSourceRemote;
							}
							for (var i=0; i < funcs.length; i++) {
								funcs[i](arg);
							}
						}
					}
				}
			}
		};
		GenericEventRegistration.getGenericEventRegistration=function () {
			if (!GenericEventRegistration.s_genericEventRegistration) {
				GenericEventRegistration.s_genericEventRegistration=new GenericEventRegistration();
			}
			return GenericEventRegistration.s_genericEventRegistration;
		};
		GenericEventRegistration.richApiMessageEventCategory=65536;
		return GenericEventRegistration;
	}());
	function _testSetRichApiMessageReadyImpl(impl) {
		GenericEventRegistration._testReadyImpl=impl;
	}
	OfficeExtension._testSetRichApiMessageReadyImpl=_testSetRichApiMessageReadyImpl;
	function _testTriggerRichApiMessageEvent(msg) {
		GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
	}
	OfficeExtension._testTriggerRichApiMessageEvent=_testTriggerRichApiMessageEvent;
	var GenericEventHandlers=(function (_super) {
		__extends(GenericEventHandlers, _super);
		function GenericEventHandlers(context, parentObject, name, eventInfo) {
			_super.call(this, context, parentObject, name, eventInfo);
			this.m_genericEventInfo=eventInfo;
		}
		GenericEventHandlers.prototype.add=function (handler) {
			var _this=this;
			if ((this._handlers.length==0) && this.m_genericEventInfo.registerFunc) {
				this.m_genericEventInfo.registerFunc();
			}
			if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
				this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				_this._handlers.push(handler);
				if (_this._handlers.length==1) {
					GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
			return new OfficeExtension.EventHandlerResult(this._context, this, handler);
		};
		GenericEventHandlers.prototype.remove=function (handler) {
			var _this=this;
			if ((this._handlers.length==1) && this.m_genericEventInfo.unregisterFunc) {
				this.m_genericEventInfo.unregisterFunc();
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				var handlers=_this._handlers;
				for (var index=handlers.length - 1; index >=0; index--) {
					if (handlers[index]===handler) {
						handlers.splice(index, 1);
						break;
					}
				}
				if (handlers.length==0) {
					GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
		};
		GenericEventHandlers.prototype.removeAll=function () {
		};
		return GenericEventHandlers;
	}(OfficeExtension.EventHandlers));
	OfficeExtension.GenericEventHandlers=GenericEventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpRequestExecutor=(function () {
		function HttpRequestExecutor() {
		}
		HttpRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			var url=requestMessage.Url;
			if (url.charAt(url.length - 1) !="/") {
				url=url+"/";
			}
			url=url+OfficeExtension.Constants.processQuery;
			url=url+"?"+OfficeExtension.Constants.flags+"="+requestFlags.toString();
			var requestInfo={
				method: "POST",
				url: url,
				headers: {},
				body: requestMessageText
			};
			requestInfo.headers[OfficeExtension.Constants.sourceLibHeader]=HttpRequestExecutor.SourceLibHeaderValue;
			requestInfo.headers["CONTENT-TYPE"]="application/json";
			if (requestMessage.Headers) {
				for (var key in requestMessage.Headers) {
					requestInfo.headers[key]=requestMessage.Headers[key];
				}
			}
			return OfficeExtension.HttpUtility.sendRequest(requestInfo)
				.then(function (responseInfo) {
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: JSON.parse(responseInfo.body) };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		HttpRequestExecutor.SourceLibHeaderValue="officejs-rest";
		return HttpRequestExecutor;
	}());
	OfficeExtension.HttpRequestExecutor=HttpRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpUtility=(function () {
		function HttpUtility() {
		}
		HttpUtility.setCustomSendRequestFunc=function (func) {
			HttpUtility.s_customSendRequestFunc=func;
		};
		HttpUtility.xhrSendRequestFunc=function (request) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var xhr=new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload=function () {
					var resp={
						statusCode: xhr.status,
						headers: OfficeExtension.Utility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
						body: xhr.responseText
					};
					resolve(resp);
				};
				xhr.onerror=function () {
					reject(new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.connectionFailure,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, xhr.statusText)
					}));
				};
				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}
				xhr.send(request.body);
			});
		};
		HttpUtility.sendRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func=HttpUtility.s_customSendRequestFunc;
			if (!func) {
				func=HttpUtility.xhrSendRequestFunc;
			}
			return func(request);
		};
		HttpUtility.setCustomSendLocalDocumentRequestFunc=function (func) {
			HttpUtility.s_customSendLocalDocumentRequestFunc=func;
		};
		HttpUtility.sendLocalDocumentRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func;
			func=HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
			return func(request);
		};
		HttpUtility.officeJsSendLocalDocumentRequestFunc=function (request) {
			request=OfficeExtension.Utility._validateLocalDocumentRequest(request);
			var requestSafeArray=OfficeExtension.Utility._buildRequestMessageSafeArray(request);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
					var response;
					if (asyncResult.status=="succeeded") {
						response=							{
								statusCode: OfficeExtension.RichApiMessageUtility.getResponseStatusCode(asyncResult),
								headers: OfficeExtension.RichApiMessageUtility.getResponseHeaders(asyncResult),
								body: OfficeExtension.RichApiMessageUtility.getResponseBody(asyncResult)
							};
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
					}
					OfficeExtension.Utility.log(JSON.stringify(response));
					resolve(response);
				});
			});
		};
		HttpUtility.validateAndNormalizeRequest=function (request) {
			if (OfficeExtension.Utility.isNullOrUndefined(request)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(request.method)) {
				request.method="GET";
			}
			request.method=request.method.toUpperCase();
		};
		HttpUtility.logRequest=function (request) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Request---");
				OfficeExtension.Utility.log(request.method+" "+request.url);
				if (request.headers) {
					for (var key in request.headers) {
						OfficeExtension.Utility.log(key+": "+request.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(request.body);
				}
			}
		};
		HttpUtility.logResponse=function (response) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Response---");
				OfficeExtension.Utility.log(""+response.statusCode);
				if (response.headers) {
					for (var key in response.headers) {
						OfficeExtension.Utility.log(key+": "+response.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(response.body);
				}
			}
		};
		HttpUtility._logBody=false;
		return HttpUtility;
	}());
	OfficeExtension.HttpUtility=HttpUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._handleIdResult(value);
		};
		return InstantiateActionResultHandler;
	}());
	OfficeExtension.InstantiateActionResultHandler=InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HostBridgeRequestExecutor=(function () {
		function HostBridgeRequestExecutor(session) {
			this.m_session=session;
		}
		HostBridgeRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var httpRequestInfo={
				url: OfficeExtension.Constants.processQuery,
				method: "POST",
				headers: requestMessage.Headers,
				body: requestMessage.Body
			};
			var message={
				id: HostBridgeSession.nextId(),
				type: 1,
				flags: requestFlags,
				message: httpRequestInfo
			};
			OfficeExtension.Utility.log(JSON.stringify(message));
			return this.m_session.sendMessageToHost(message)
				.then(function (nativeBridgeResponse) {
				OfficeExtension.Utility.log("Received response: "+JSON.stringify(nativeBridgeResponse));
				var responseInfo=nativeBridgeResponse.message;
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: responseInfo.body };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		return HostBridgeRequestExecutor;
	}());
	var HostBridgeSession=(function (_super) {
		__extends(HostBridgeSession, _super);
		function HostBridgeSession(bridge) {
			var _this=this;
			_super.call(this);
			this.m_promiseResolver={};
			this.m_bridge=bridge;
			this.m_bridge.onMessageFromHost=function (msg) {
				_this.onMessageFromHost(msg);
			};
		}
		HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		HostBridgeSession.prototype._createRequestExecutorOrNull=function () {
			OfficeExtension.Utility.log("NativeBridgeSession::CreateRequestExecutor");
			return new HostBridgeRequestExecutor(this);
		};
		Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		HostBridgeSession.init=function (bridge) {
			if (bridge && typeof (bridge)==="object") {
				var session=new HostBridgeSession(bridge);
				OfficeExtension.ClientRequestContext._overrideSession=session;
				OfficeExtension.HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
					var bridgeMessage={
						id: HostBridgeSession.nextId(),
						type: 1,
						flags: 0,
						message: request
					};
					return session.sendMessageToHost(bridgeMessage)
						.then(function (bridgeResponse) {
						var responseInfo=bridgeResponse.message;
						return responseInfo;
					});
				});
			}
		};
		HostBridgeSession.prototype.sendMessageToHost=function (message) {
			var _this=this;
			this.m_bridge.sendMessageToHost(JSON.stringify(message));
			var ret=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_promiseResolver[message.id]=resolve;
			});
			return ret;
		};
		HostBridgeSession.prototype.onMessageFromHost=function (messageText) {
			if (messageText==="test") {
				if (HostBridgeTest._testFunc) {
					HostBridgeTest._testFunc();
				}
			}
			else {
				var message=JSON.parse(messageText);
				if (typeof (message.id)==="number") {
					var resolve=this.m_promiseResolver[message.id];
					if (resolve) {
						resolve(message);
					}
					delete this.m_promiseResolver[message.id];
				}
			}
		};
		HostBridgeSession.nextId=function () {
			return HostBridgeSession.s_nextId++;
		};
		HostBridgeSession.s_nextId=1;
		return HostBridgeSession;
	}(OfficeExtension.SessionBase));
	var HostBridge=(function () {
		function HostBridge() {
		}
		HostBridge.init=function (bridge) {
			HostBridgeSession.init(bridge);
		};
		return HostBridge;
	}());
	OfficeExtension.HostBridge=HostBridge;
	if (typeof (_richApiNativeBridge)==="object" && _richApiNativeBridge) {
		HostBridge.init(_richApiNativeBridge);
	}
	var HostBridgeTest=(function () {
		function HostBridgeTest() {
		}
		HostBridgeTest.setTestFunc=function (func) {
			HostBridgeTest._testFunc=func;
		};
		return HostBridgeTest;
	}());
	OfficeExtension.HostBridgeTest=HostBridgeTest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isWriteOperation=false;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			set: function (value) {
				this.m_isWriteOperation=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			set: function (value) {
				this.m_isRestrictedResourceAccess=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isCollection", {
			get: function () {
				return this.m_isCollection;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
			get: function () {
				return this.m_isInvalidAfterRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
			get: function () {
				return this.m_parentObjectPath;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
			get: function () {
				return this.m_argumentObjectPaths;
			},
			set: function (value) {
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
				if (!value &&
					this.m_objectPathInfo.ObjectPathType===6 &&
					this.m_savedObjectPathInfo) {
					ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
					this.m_parentObjectPath=this.m_savedObjectPathInfo.parent;
					this.m_isValid=true;
					this.m_savedObjectPathInfo=null;
				}
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
			get: function () {
				return this.m_originalObjectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
			get: function () {
				return this.m_getByIdMethodName;
			},
			set: function (value) {
				this.m_getByIdMethodName=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.resetForUpdateUsingObjectData();
			this.m_objectPathInfo.ObjectPathType=7;
			this.m_objectPathInfo.Name="";
			this.m_parentObjectPath=null;
		};
		ObjectPath.prototype.saveOriginalObjectPathInfo=function () {
			if (OfficeExtension.config.extendedErrorLogging && !this.m_originalObjectPathInfo) {
				this.m_originalObjectPathInfo={};
				ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo);
			}
		};
		ObjectPath.prototype.updateUsingObjectData=function (value, clientObject) {
			var referenceId=value[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				if (!this.m_savedObjectPathInfo &&
					!this.isInvalidAfterRequest &&
					ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
					var pathInfo={};
					ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
					this.m_savedObjectPathInfo={
						pathInfo: pathInfo,
						parent: this.m_parentObjectPath
					};
				}
				this.saveOriginalObjectPathInfo();
				this.resetForUpdateUsingObjectData();
				this.m_objectPathInfo.ObjectPathType=6;
				this.m_objectPathInfo.Name=referenceId;
				delete this.m_objectPathInfo.ParentObjectPathId;
				this.m_parentObjectPath=null;
				return;
			}
			var collectionPropertyPath=clientObject[OfficeExtension.Constants.collectionPropertyPath];
			if (!OfficeExtension.Utility.isNullOrEmptyString(collectionPropertyPath)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					var propNames=collectionPropertyPath.split(".");
					var parent_1=clientObject.context[propNames[0]];
					for (var i=1; i < propNames.length; i++) {
						parent_1=parent_1[propNames[i]];
					}
					this.saveOriginalObjectPathInfo();
					this.resetForUpdateUsingObjectData();
					this.m_parentObjectPath=parent_1._objectPath;
					this.m_objectPathInfo.ParentObjectPathId=this.m_parentObjectPath.objectPathInfo.Id;
					this.m_objectPathInfo.ObjectPathType=5;
					this.m_objectPathInfo.Name="";
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
			var parentIsCollection=this.parentObjectPath && this.parentObjectPath.isCollection;
			var getByIdMethodName=this.getByIdMethodName;
			if (parentIsCollection || !OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					this.saveOriginalObjectPathInfo();
					this.resetForUpdateUsingObjectData();
					if (!OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
						this.m_objectPathInfo.ObjectPathType=3;
						this.m_objectPathInfo.Name=getByIdMethodName;
						this.m_getByIdMethodName=null;
					}
					else {
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name="";
					}
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
		};
		ObjectPath.prototype.resetForUpdateUsingObjectData=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_isWriteOperation=false;
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.isRestorableObjectPath=function (objectPathType) {
			return (objectPathType===1 ||
				objectPathType===5 ||
				objectPathType===3 ||
				objectPathType===4);
		};
		ObjectPath.copyObjectPathInfo=function (src, dest) {
			dest.Id=src.Id;
			dest.ArgumentInfo=src.ArgumentInfo;
			dest.Name=src.Name;
			dest.ObjectPathType=src.ObjectPathType;
			dest.ParentObjectPathId=src.ParentObjectPathId;
		};
		return ObjectPath;
	}());
	OfficeExtension.ObjectPath=ObjectPath;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 1, Name: "" };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection, isRestrictedResourceAccess) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 2, Name: typeName };
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.isWriteOperation=(operationType !=1);
			ret.getByIdMethodName=getByIdMethodName;
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createReferenceIdObjectPath=function (context, referenceId) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 6,
				Name: referenceId,
				ArgumentInfo: {}
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			var objectPathInfo=objectPathInfo=				{
					Id: context._nextId(),
					ObjectPathType: 5,
					Name: "",
					ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[OfficeExtension.Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: OfficeExtension.Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		return ObjectPathFactory;
	}());
	OfficeExtension.ObjectPathFactory=ObjectPathFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor(context) {
			this.m_context=context;
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (result.status=="succeeded") {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
						_this.m_context._processOfficeJsErrorResponse(result.error.code, response);
					}
					resolve(response);
				});
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue="officejs";
		return OfficeJsRequestExecutor;
	}());
	OfficeExtension.OfficeJsRequestExecutor=OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		_Internal.OfficeRequire=function () {
			return null;
		}();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				return (function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' &&
						typeof importScripts !=='undefined' &&
						typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=setImmediate;
						}
						return function () {
							nextTick(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useVertxTimer() {
						return function () {
							lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useMutationObserver() {
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() { }
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$selfFullfillment() {
						return new TypeError("You cannot resolve a promise with itself");
					}
					function lib$es6$promise$$internal$$cannotReturnOwn() {
						return new TypeError('A promises callback cannot return that same promise.');
					}
					function lib$es6$promise$$internal$$getThen(promise) {
						try {
							return promise.then;
						}
						catch (error) {
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
							return lib$es6$promise$$internal$$GET_THEN_ERROR;
						}
					}
					function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
						try {
							then.call(value, fulfillmentHandler, rejectionHandler);
						}
						catch (e) {
							return e;
						}
					}
					function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
						lib$es6$promise$asap$$asap(function (promise) {
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, thenable._result);
						}
						else {
							lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function (reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
					}
					function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
							else if (lib$es6$promise$utils$$isFunction(then)) {
								lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
							}
							else {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
						}
					}
					function lib$es6$promise$$internal$$resolve(promise, value) {
						if (promise===value) {
							lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
						}
						else if (lib$es6$promise$utils$$objectOrFunction(value)) {
							lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
						}
						else {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
					}
					function lib$es6$promise$$internal$$publishRejection(promise) {
						if (promise._onerror) {
							promise._onerror(promise._result);
						}
						lib$es6$promise$$internal$$publish(promise);
					}
					function lib$es6$promise$$internal$$fulfill(promise, value) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, value);
						}
					}
					function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
						try {
							resolver(function resolvePromise(value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function rejectPromise(reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
						catch (e) {
							lib$es6$promise$$internal$$reject(promise, e);
						}
					}
					function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
								lib$es6$promise$asap$$asap(function () {
									lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
								});
							}
							else {
								lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
							}
							return child;
						},
						'catch': function (onRejection) {
							return this.then(null, onRejection);
						}
					};
					return lib$es6$promise$promise$$default;
				}).call(this);
			}
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		function isEdgeLessThan14() {
			var userAgent=window.navigator.userAgent;
			var versionIdx=userAgent.indexOf("Edge/");
			if (versionIdx >=0) {
				userAgent=userAgent.substring(versionIdx+5, userAgent.length);
				if (userAgent < "14.14393")
					return true;
				else
					return false;
			}
			return false;
		}
		function determinePromise() {
			if (typeof (window)==="undefined" && typeof (Promise)==="function") {
				return Promise;
			}
			if (typeof (window) !=="undefined" && window.Promise) {
				if (isEdgeLessThan14()) {
					return _Internal.PromiseImpl.Init();
				}
				else {
					return window.Promise;
				}
			}
			else {
				return _Internal.PromiseImpl.Init();
			}
		}
		_Internal.OfficePromise=determinePromise();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var OfficePromise=_Internal.OfficePromise;
	OfficeExtension.Promise=OfficePromise;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function (object, resultValue) {
			var shouldAutoTrack=(this.m_context._autoCleanup &&
				!object[OfficeExtension.Constants.isTracked] &&
				object !==this.m_context._rootObject &&
				resultValue &&
				!OfficeExtension.Utility.isNullOrEmptyString(resultValue[OfficeExtension.Constants.referenceId]));
			if (shouldAutoTrack) {
				this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			if (object[OfficeExtension.Constants.isTracked]) {
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				return;
			}
			var referenceId=object[OfficeExtension.Constants.referenceId];
			var donotKeepReference=object._objectPath.objectPathInfo[OfficeExtension.Constants.objectPathInfoDoNotKeepReferenceFieldName];
			if (donotKeepReference) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.objectIsUntracked), null);
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			object._objectPath.objectPathInfo[OfficeExtension.Constants.objectPathInfoDoNotKeepReferenceFieldName]=true;
			object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
			}
			delete object[OfficeExtension.Constants.isTracked];
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	}());
	OfficeExtension.TrackedObjects=TrackedObjects;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RequestPrettyPrinter=(function () {
		function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
			if (!globalObjName) {
				globalObjName="root";
			}
			this.m_globalObjName=globalObjName;
			this.m_referencedObjectPaths=referencedObjectPaths;
			this.m_actions=actions;
			this.m_statements=[];
			this.m_variableNameForObjectPathMap={};
			this.m_variableNameToObjectPathMap={};
			this.m_declaredObjectPathMap={};
			this.m_showDispose=showDispose;
			this.m_removePII=removePII;
		}
		RequestPrettyPrinter.prototype.process=function () {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
			}
			return this.m_statements;
		};
		RequestPrettyPrinter.prototype.processForDebugStatementInfo=function (actionIndex) {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var surroundingCount=5;
			this.m_statements=[];
			var oneStatement="";
			var statementIndex=-1;
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
				if (actionIndex==i) {
					statementIndex=this.m_statements.length - 1;
				}
				if (statementIndex >=0 && this.m_statements.length > statementIndex+surroundingCount+1) {
					break;
				}
			}
			if (statementIndex < 0) {
				return null;
			}
			var startIndex=statementIndex - surroundingCount;
			if (startIndex < 0) {
				startIndex=0;
			}
			var endIndex=statementIndex+1+surroundingCount;
			if (endIndex > this.m_statements.length) {
				endIndex=this.m_statements.length;
			}
			var surroundingStatements=[];
			if (startIndex !=0) {
				surroundingStatements.push("...");
			}
			for (var i_1=startIndex; i_1 < statementIndex; i_1++) {
				surroundingStatements.push(this.m_statements[i_1]);
			}
			surroundingStatements.push("// >>>>>");
			surroundingStatements.push(this.m_statements[statementIndex]);
			surroundingStatements.push("// <<<<<");
			for (var i_2=statementIndex+1; i_2 < endIndex; i_2++) {
				surroundingStatements.push(this.m_statements[i_2]);
			}
			if (endIndex < this.m_statements.length) {
				surroundingStatements.push("...");
			}
			return {
				statement: this.m_statements[statementIndex],
				surroundingStatements: surroundingStatements
			};
		};
		RequestPrettyPrinter.prototype.processOneAction=function (action) {
			var actionInfo=action.actionInfo;
			switch (actionInfo.ActionType) {
				case 1:
					this.processInstantiateAction(action);
					break;
				case 3:
					this.processMethodAction(action);
					break;
				case 2:
					this.processQueryAction(action);
					break;
				case 7:
					this.processQueryAsJsonAction(action);
					break;
				case 6:
					this.processRecursiveQueryAction(action);
					break;
				case 4:
					this.processSetPropertyAction(action);
					break;
				case 5:
					this.processTraceAction(action);
					break;
				case 8:
					this.processEnsureUnchangedAction(action);
					break;
				case 9:
					this.processUpdateAction(action);
					break;
			}
		};
		RequestPrettyPrinter.prototype.processInstantiateAction=function (action) {
			var objId=action.actionInfo.ObjectPathId;
			var objPath=this.m_referencedObjectPaths[objId];
			var varName=this.getObjVarName(objId);
			if (!this.m_declaredObjectPathMap[objId]) {
				var statement="var "+varName+"="+this.buildObjectPathExpressionWithParent(objPath)+";";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
				this.m_declaredObjectPathMap[objId]=varName;
			}
			else {
				var statement="// Instantiate {"+varName+"}";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
			}
		};
		RequestPrettyPrinter.prototype.processMethodAction=function (action) {
			var methodName=action.actionInfo.Name;
			if (methodName==="_KeepReference") {
				if (!OfficeExtension._internalConfig.showInternalApiInDebugInfo) {
					return;
				}
				methodName="track";
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(methodName)+"("+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".load("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAsJsonAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".retrieve("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processRecursiveQueryAction=function (action) {
			var queryExp="";
			if (action.actionInfo.RecursiveQueryInfo) {
				queryExp=JSON.stringify(action.actionInfo.RecursiveQueryInfo);
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".loadRecursive("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processSetPropertyAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(action.actionInfo.Name)+"="+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+";";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processTraceAction=function (action) {
			var statement="context.trace();";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processEnsureUnchangedAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".ensureUnchanged("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processUpdateAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".update("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant=function (statement, action) {
			var _this=this;
			if (this.m_showDispose) {
				var lastUsedObjectPathIds=action.actionInfo.L;
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					var objectNamesToDispose=lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(", ");
					return statement+" // And then dispose {"+objectNamesToDispose+"}";
				}
			}
			return statement;
		};
		RequestPrettyPrinter.prototype.buildQueryExpression=function (action) {
			if (action.actionInfo.QueryInfo) {
				var option={};
				option.select=action.actionInfo.QueryInfo.Select;
				option.expand=action.actionInfo.QueryInfo.Expand;
				option.skip=action.actionInfo.QueryInfo.Skip;
				option.top=action.actionInfo.QueryInfo.Top;
				if (typeof (option.top)==="undefined" && typeof (option.skip)==="undefined" && typeof (option.expand)==="undefined") {
					if (typeof (option.select)==="undefined") {
						return "";
					}
					else {
						return JSON.stringify(option.select);
					}
				}
				else {
					return JSON.stringify(option);
				}
			}
			return "";
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent=function (objPath) {
			var hasParent=objPath.objectPathInfo.ObjectPathType==5 ||
				objPath.objectPathInfo.ObjectPathType==3 ||
				objPath.objectPathInfo.ObjectPathType==4;
			if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
				return this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId)+"."+this.buildObjectPathExpression(objPath);
			}
			return this.buildObjectPathExpression(objPath);
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpression=function (objPath) {
			var expr=this.buildObjectPathInfoExpression(objPath.objectPathInfo);
			var originalObjectPathInfo=objPath.originalObjectPathInfo;
			if (originalObjectPathInfo) {
				expr=expr+" /* originally "+this.buildObjectPathInfoExpression(originalObjectPathInfo)+" */";
			}
			return expr;
		};
		RequestPrettyPrinter.prototype.buildObjectPathInfoExpression=function (objectPathInfo) {
			switch (objectPathInfo.ObjectPathType) {
				case 1:
					return "context."+this.m_globalObjName;
				case 5:
					return "getItem("+this.buildArgumentsExpression(objectPathInfo.ArgumentInfo)+")";
				case 3:
					return OfficeExtension.Utility._toCamelLowerCase(objectPathInfo.Name)+"("+this.buildArgumentsExpression(objectPathInfo.ArgumentInfo)+")";
				case 2:
					return objectPathInfo.Name+".newObject()";
				case 7:
					return "null";
				case 4:
					return OfficeExtension.Utility._toCamelLowerCase(objectPathInfo.Name);
				case 6:
					return "context."+this.m_globalObjName+"._getObjectByReferenceId("+JSON.stringify(objectPathInfo.Name)+")";
			}
		};
		RequestPrettyPrinter.prototype.buildArgumentsExpression=function (args) {
			var ret="";
			if (!args.Arguments || args.Arguments.length===0) {
				return ret;
			}
			if (this.m_removePII) {
				if (typeof (args.Arguments[0])==="undefined") {
					return ret;
				}
				return "...";
			}
			for (var i=0; i < args.Arguments.length; i++) {
				if (i > 0) {
					ret=ret+", ";
				}
				ret=ret+this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
			}
			if (ret==="undefined") {
				ret="";
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.buildArgumentLiteral=function (value, objectPathId) {
			if (typeof value=="number" && value===objectPathId) {
				return this.getObjVarName(objectPathId);
			}
			else {
				return JSON.stringify(value);
			}
		};
		RequestPrettyPrinter.prototype.getObjVarNameBase=function (objectPathId) {
			var ret="v";
			var objPath=this.m_referencedObjectPaths[objectPathId];
			if (objPath) {
				switch (objPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=this.m_globalObjName;
						break;
					case 4:
						ret=OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
						break;
					case 3:
						var methodName=objPath.objectPathInfo.Name;
						if (methodName.length > 3 && methodName.substr(0, 3)==="Get") {
							methodName=methodName.substr(3);
						}
						ret=OfficeExtension.Utility._toCamelLowerCase(methodName);
						break;
					case 5:
						var parentName=this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
						if (parentName.charAt(parentName.length - 1)==="s") {
							ret=parentName.substr(0, parentName.length - 1);
						}
						else {
							ret=parentName+"Item";
						}
						break;
				}
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.getObjVarName=function (objectPathId) {
			if (this.m_variableNameForObjectPathMap[objectPathId]) {
				return this.m_variableNameForObjectPathMap[objectPathId];
			}
			var ret=this.getObjVarNameBase(objectPathId);
			if (!this.m_variableNameToObjectPathMap[ret]) {
				this.m_variableNameForObjectPathMap[objectPathId]=ret;
				this.m_variableNameToObjectPathMap[ret]=objectPathId;
				return ret;
			}
			var i=1;
			while (this.m_variableNameToObjectPathMap[ret+i.toString()]) {
				i++;
			}
			ret=ret+i.toString();
			this.m_variableNameForObjectPathMap[objectPathId]=ret;
			this.m_variableNameToObjectPathMap[ret]=objectPathId;
			return ret;
		};
		return RequestPrettyPrinter;
	}());
	OfficeExtension.RequestPrettyPrinter=RequestPrettyPrinter;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStrings=(function () {
		function ResourceStrings() {
		}
		ResourceStrings.cannotRegisterEvent="CannotRegisterEvent";
		ResourceStrings.connectionFailureWithStatus="ConnectionFailureWithStatus";
		ResourceStrings.connectionFailureWithDetails="ConnectionFailureWithDetails";
		ResourceStrings.invalidObjectPath="InvalidObjectPath";
		ResourceStrings.invalidRequestContext="InvalidRequestContext";
		ResourceStrings.invalidArgument="InvalidArgument";
		ResourceStrings.invalidArgumentGeneric="InvalidArgumentGeneric";
		ResourceStrings.propertyNotLoaded="PropertyNotLoaded";
		ResourceStrings.runMustReturnPromise="RunMustReturnPromise";
		ResourceStrings.timeout="Timeout";
		ResourceStrings.propertyDoesNotExist="PropertyDoesNotExist";
		ResourceStrings.attemptingToSetReadOnlyProperty="AttemptingToSetReadOnlyProperty";
		ResourceStrings.moreInfoInnerError="MoreInfoInnerError";
		ResourceStrings.cannotApplyPropertyThroughSetMethod="CannotApplyPropertyThroughSetMethod";
		ResourceStrings.valueNotLoaded="ValueNotLoaded";
		ResourceStrings.invalidOrTimedOutSessionMessage="InvalidOrTimedOutSessionMessage";
		ResourceStrings.invalidOperationInCellEditMode="InvalidOperationInCellEditMode";
		ResourceStrings.objectIsUntracked="ObjectIsUntracked";
		ResourceStrings.customFunctionDefintionMissing="CustomFunctionDefintionMissing";
		ResourceStrings.customFunctionImplementationMissing="CustomFunctionImplementationMissing";
		ResourceStrings.customFunctionNameContainsBadChars="CustomFunctionNameContainsBadChars";
		ResourceStrings.customFunctionNameCannotSplit="CustomFunctionNameCannotSplit";
		ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch="CustomFunctionUnexpectedNumberOfEntriesInResultBatch";
		ResourceStrings.customFunctionCancellationHandlerMissing="CustomFunctionCancellationHandlerMissing";
		ResourceStrings.apiNotFoundDetails="ApiNotFoundDetails";
		ResourceStrings.pendingBatchInProgress="PendingBatchInProgress";
		ResourceStrings.notInsideBatch="NotInsideBatch";
		ResourceStrings.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ResourceStrings;
	}());
	OfficeExtension.ResourceStrings=ResourceStrings;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStringValues=(function () {
		function ResourceStringValues() {
		}
		ResourceStringValues.CannotRegisterEvent="The event handler cannot be registered.";
		ResourceStringValues.ConnectionFailureWithStatus="The request failed with status code of {0}.";
		ResourceStringValues.ConnectionFailureWithDetails="The request failed with status code of {0}, error code {1} and the following error message: {2}";
		ResourceStringValues.InvalidArgument="The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.";
		ResourceStringValues.InvalidObjectPath="The object path '{0}' isn't working for what you're trying to do. If you're using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object's lifetime.";
		ResourceStringValues.InvalidRequestContext="Cannot use the object across different request contexts.";
		ResourceStringValues.PropertyNotLoaded="The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.";
		ResourceStringValues.RunMustReturnPromise="The batch function passed to the \".run\" method didn't return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".";
		ResourceStringValues.Timeout="The operation has timed out.";
		ResourceStringValues.ValueNotLoaded="The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.";
		ResourceStringValues.InvalidOrTimedOutSessionMessage="Your Office Online session has expired or is invalid. To continue, refresh the page.";
		ResourceStringValues.InvalidOperationInCellEditMode="Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.";
		ResourceStringValues.CustomFunctionDefintionMissing="A property with this name that represents the function's definition must exist on Excel.CustomFunctions.";
		ResourceStringValues.CustomFunctionImplementationMissing="The property with this name on Excel.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.";
		ResourceStringValues.CustomFunctionNameContainsBadChars="The function name may only contain letters, digits, underscores, and periods.";
		ResourceStringValues.CustomFunctionNameCannotSplit="The function name must contain a non-empty namespace and a non-empty short name.";
		ResourceStringValues.CustomFunctionUnexpectedNumberOfEntriesInResultBatch="The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.";
		ResourceStringValues.CustomFunctionCancellationHandlerMissing="The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.";
		ResourceStringValues.ApiNotFoundDetails="The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.";
		ResourceStringValues.PendingBatchInProgress="There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.";
		ResourceStringValues.NotInsideBatch="Operations may not be invoked outside of a batch method.";
		ResourceStringValues.CannotUpdateReadOnlyProperty="The property '{0}' is read-only and it cannot be updated.";
		ResourceStringValues.ObjectIsUntracked="The object is untracked.";
		return ResourceStringValues;
	}());
	OfficeExtension.ResourceStringValues=ResourceStringValues;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			OfficeExtension.Utility.log("Request:");
			OfficeExtension.Utility.log(requestMessageText);
			var headers={};
			headers[OfficeExtension.Constants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.buildResponseOnSuccess=function (responseBody, responseHeaders) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			return response;
		};
		RichApiMessageUtility.buildResponseOnError=function (errorCode, message) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=OfficeExtension.ErrorCodes.generalException;
			response.ErrorMessage=message;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
				response.ErrorCode=OfficeExtension.ErrorCodes.invalidOrTimedOutSession;
				response.ErrorMessage=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOrTimedOutSessionMessage);
			}
			return response;
		};
		RichApiMessageUtility.buildHttpResponseFromOfficeJsError=function (errorCode, message) {
			var statusCode=500;
			var errorBody={};
			errorBody["error"]={};
			errorBody["error"]["code"]=OfficeExtension.ErrorCodes.generalException;
			errorBody["error"]["message"]=message;
			if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				statusCode=403;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				statusCode=429;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId="";
			var instanceId="";
			var marketplaceType="";
			return [
				customData,
				method,
				path,
				headerArray,
				body,
				appPermission,
				requestFlags,
				solutionId,
				instanceId,
				marketplaceType
			];
		};
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2];
			if (typeof (ret)==="string") {
				return ret;
			}
			var arr=ret;
			return arr.join("");
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		return RichApiMessageUtility;
	}());
	OfficeExtension.RichApiMessageUtility=RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Utility=(function () {
		function Utility() {
		}
		Utility.checkArgumentNull=function (value, name) {
			if (Utility.isNullOrUndefined(value)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: name });
			}
		};
		Utility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isUndefined=function (value) {
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		Utility.isPlainJsonObject=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return false;
			}
			if (typeof (value) !=="object") {
				return false;
			}
			return Object.getPrototypeOf(value)===Object.getPrototypeOf({});
		};
		Utility.trim=function (str) {
			return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
		};
		Utility.caseInsensitiveCompareString=function (str1, str2) {
			if (Utility.isNullOrUndefined(str1)) {
				return Utility.isNullOrUndefined(str2);
			}
			else {
				if (Utility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		Utility.adjustToDateTime=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return null;
			}
			if (typeof (value)==="string") {
				return new Date(value);
			}
			if (Array.isArray(value)) {
				var arr=value;
				for (var i=0; i < arr.length; i++) {
					arr[i]=Utility.adjustToDateTime(arr[i]);
				}
				return arr;
			}
			throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "date" });
		};
		Utility.isReadonlyRestRequest=function (method) {
			return Utility.caseInsensitiveCompareString(method, "GET");
		};
		Utility.setMethodArguments=function (context, argumentInfo, args) {
			if (Utility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
			}
			return referencedObjectPaths;
		};
		Utility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof OfficeExtension.ClientObject) {
					var clientObject=args[i];
					Utility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else if (Utility.isPlainJsonObject(args[i])) {
					referencedObjectPathIds.push(0);
					Utility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths);
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		Utility.replaceClientObjectPropertiesWithObjectPathIds=function (value, referencedObjectPaths) {
			for (var key in value) {
				var propValue=value[key];
				if (propValue instanceof OfficeExtension.ClientObject) {
					referencedObjectPaths.push(propValue._objectPath);
					value[key]=(_a={}, _a[OfficeExtension.Constants.objectPathIdPrivate]=propValue._objectPath.objectPathInfo.Id, _a);
				}
				else if (Array.isArray(propValue)) {
					for (var i=0; i < propValue.length; i++) {
						if (propValue[i] instanceof OfficeExtension.ClientObject) {
							var elem=propValue[i];
							referencedObjectPaths.push(elem._objectPath);
							propValue[i]=(_b={}, _b[OfficeExtension.Constants.objectPathIdPrivate]=elem._objectPath.objectPathInfo.Id, _b);
						}
						else if (Utility.isPlainJsonObject(propValue[i])) {
							Utility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths);
						}
					}
				}
				else if (Utility.isPlainJsonObject(propValue)) {
					Utility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
				}
				else {
				}
			}
			var _a, _b;
		};
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value, clientObject);
			}
		};
		Utility.tryGetObjectIdFromLoadOrRetrieveResult=function (value) {
			var id=value[OfficeExtension.Constants.id];
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idLowerCase];
			}
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idPrivate];
			}
			return id;
		};
		Utility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.invalidObjectPath,
						message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
						debugInfo: {
							errorLocation: Utility.getObjectPathExpression(objectPath)
						}
					});
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		Utility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidObjectPath,
								message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
							});
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		Utility.validateContext=function (context, obj) {
			if (obj && obj.context !==context) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidRequestContext,
					message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidRequestContext)
				});
			}
		};
		Utility.log=function (message) {
			if (Utility._logEnabled && typeof (console) !=="undefined" && console.log) {
				console.log(message);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj;
		};
		Utility.loadAndSync=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj.context.sync().then(function () { return clientObj; });
		};
		Utility.retrieve=function (clientObj, option) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!Utility.isSetSupported("RichApiRuntime", "1.1");
			}
			var result=new OfficeExtension.RetrieveResultImpl(clientObj, shouldPolyfill);
			var queryOption=OfficeExtension.ClientRequestContext._parseQueryOption(option);
			var action;
			if (shouldPolyfill) {
				action=OfficeExtension.ActionFactory.createQueryAction(clientObj.context, clientObj, queryOption);
			}
			else {
				action=OfficeExtension.ActionFactory.createQueryAsJsonAction(clientObj.context, clientObj, queryOption);
			}
			clientObj.context._pendingRequest.addActionResultHandler(action, result);
			return result;
		};
		Utility.retrieveAndSync=function (clientObj, option) {
			var result=Utility.retrieve(clientObj, option);
			return clientObj.context.sync().then(function () { return result; });
		};
		Utility.isSetSupported=function (apiSetName, apiSetVersion) {
			if (typeof (window) !=="undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
				return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
			}
			return true;
		};
		Utility._parseSelectExpand=function (select) {
			var args=[];
			if (!Utility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(",");
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=sanitizeForAnyItemsSlash(propertyName.trim());
					if (propertyName.length > 0) {
						args.push(propertyName);
					}
				}
			}
			return args;
			function sanitizeForAnyItemsSlash(propertyName) {
				var propertyNameLower=propertyName.toLowerCase();
				if (propertyNameLower==="items" || propertyNameLower==="items/") {
					return '*';
				}
				var itemsSlashLength=6;
				var isItemsSlashOrItemsDot=propertyNameLower.substr(0, itemsSlashLength)==="items/" ||
					propertyNameLower.substr(0, itemsSlashLength)==="items.";
				if (isItemsSlashOrItemsDot) {
					propertyName=propertyName.substr(itemsSlashLength);
				}
				return propertyName.replace(new RegExp("[\/\.]items[\/\.]", "gi"), "/");
			}
		};
		Utility.toJson=function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
			var result={};
			for (var prop in scalarProperties) {
				var value=scalarProperties[prop];
				if (typeof value !=="undefined") {
					result[prop]=value;
				}
			}
			for (var prop in navigationProperties) {
				var value=navigationProperties[prop];
				if (typeof value !=="undefined") {
					if (value[Utility.fieldName_isCollection] && (typeof value[Utility.fieldName_m__items] !=="undefined")) {
						result[prop]=value.toJSON()["items"];
					}
					else {
						result[prop]=value.toJSON();
					}
				}
			}
			if (collectionItemsIfAny) {
				result["items"]=collectionItemsIfAny.map(function (item) { return item.toJSON(); });
			}
			return result;
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new OfficeExtension._Internal.RuntimeError({
				code: resourceId,
				message: Utility._getResourceString(resourceId, arg),
				debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
			});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return (new OfficeExtension._Internal.RuntimeError({
				code: code,
				message: message,
				debugInfo: { errorLocation: location }
			}));
		};
		Utility._getResourceString=function (resourceId, arg) {
			var ret;
			if (typeof (window) !=="undefined" && window.Strings && window.Strings.OfficeOM) {
				var stringName="L_"+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!ret) {
				ret=OfficeExtension.ResourceStringValues[resourceId];
			}
			if (!ret) {
				ret=resourceId;
			}
			if (!Utility.isNullOrUndefined(arg)) {
				if (Array.isArray(arg)) {
					var arrArg=arg;
					ret=Utility._formatString(ret, arrArg);
				}
				else {
					ret=ret.replace("{0}", arg);
				}
			}
			return ret;
		};
		Utility._formatString=function (format, arrArg) {
			return format.replace(/\{\d\}/g, function (v) {
				var position=parseInt(v.substr(1, v.length - 2));
				if (position < arrArg.length) {
					return arrArg[position];
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({ argumentName: "format" });
				}
			});
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				throw Utility.createPropertyNotLoadedException(entityName, propertyName);
			}
		};
		Utility.createPropertyNotLoadedException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.propertyNotLoaded,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.createCannotUpdateReadOnlyPropertyException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.cannotUpdateReadOnlyProperty,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.throwIfApiNotSupported=function (apiFullName, apiSetName, apiSetVersion, hostName) {
			if (!Utility._doApiNotSupportedCheck) {
				return;
			}
			if (!Utility.isSetSupported(apiSetName, apiSetVersion)) {
				var message=Utility._getResourceString(OfficeExtension.ResourceStrings.apiNotFoundDetails, [apiFullName, apiSetName+" "+apiSetVersion, hostName]);
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.apiNotFound,
					message: message,
					debugInfo: { errorLocation: apiFullName }
				});
			}
		};
		Utility.getObjectPathExpression=function (objectPath) {
			var ret="";
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=ret;
						break;
					case 2:
						ret="new()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 3:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+"()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 4:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? "." : "")+ret;
						break;
					case 5:
						ret="getItem()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 6:
						ret="_reference()"+(ret.length > 0 ? "." : "")+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		Utility._createPromiseFromResult=function (value) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				resolve(value);
			});
		};
		Utility._createTimeoutPromise=function (timeout) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				setTimeout(function () {
					resolve(null);
				}, timeout);
			});
		};
		Utility.promisify=function (action) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var callback=function (result) {
					if (result.status=="failed") {
						reject(result.error);
					}
					else {
						resolve(result.value);
					}
				};
				action(callback);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!Utility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		Utility._isLocalDocumentUrl=function (url) {
			return Utility._getLocalDocumentUrlPrefixLength(url) > 0;
		};
		Utility._getLocalDocumentUrlPrefixLength=function (url) {
			var localDocumentPrefixes=["http://document.localhost", "https://document.localhost", "//document.localhost"];
			var urlLower=url.toLowerCase().trim();
			for (var i=0; i < localDocumentPrefixes.length; i++) {
				if (urlLower===localDocumentPrefixes[i]) {
					return localDocumentPrefixes[i].length;
				}
				else if (urlLower.substr(0, localDocumentPrefixes[i].length+1)===localDocumentPrefixes[i]+"/") {
					return localDocumentPrefixes[i].length+1;
				}
			}
			return 0;
		};
		Utility._validateLocalDocumentRequest=function (request) {
			var index=Utility._getLocalDocumentUrlPrefixLength(request.url);
			if (index <=0) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			var path=request.url.substr(index);
			var pathLower=path.toLowerCase();
			if (pathLower==="_api") {
				path="";
			}
			else if (pathLower.substr(0, "_api/".length)==="_api/") {
				path=path.substr("_api/".length);
			}
			return {
				method: request.method,
				url: path,
				headers: request.headers,
				body: request.body
			};
		};
		Utility._buildRequestMessageSafeArray=function (request) {
			var requestFlags=0;
			if (!Utility.isReadonlyRestRequest(request.method)) {
				requestFlags=1;
			}
			if (request.url.substr(0, OfficeExtension.Constants.processQuery.length).toLowerCase()===OfficeExtension.Constants.processQuery.toLowerCase()) {
				var index=request.url.indexOf("?");
				if (index > 0) {
					var queryString=request.url.substr(index+1);
					var parts=queryString.split("&");
					for (var i=0; i < parts.length; i++) {
						var keyvalue=parts[i].split("=");
						if (keyvalue[0].toLowerCase()===OfficeExtension.Constants.flags) {
							var flags=parseInt(keyvalue[1]);
							requestFlags=flags;
							requestFlags=requestFlags & 1;
							break;
						}
					}
				}
			}
			return OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray("", requestFlags, request.method, request.url, request.headers, request.body);
		};
		Utility._parseHttpResponseHeaders=function (allResponseHeaders) {
			var responseHeaders={};
			if (!Utility.isNullOrEmptyString(allResponseHeaders)) {
				var regex=new RegExp("\r?\n");
				var entries=allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=Utility.trim(key);
							value=Utility.trim(value);
							responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
			return responseHeaders;
		};
		Utility._parseErrorResponse=function (responseInfo) {
			var errorObj=null;
			if (Utility.isPlainJsonObject(responseInfo.body)) {
				errorObj=responseInfo.body;
			}
			else if (!Utility.isNullOrEmptyString(responseInfo.body)) {
				var errorResponseBody=Utility.trim(responseInfo.body);
				try {
					errorObj=JSON.parse(errorResponseBody);
				}
				catch (e) {
					Utility.log("Error when parse "+errorResponseBody);
				}
			}
			var errorMessage;
			var errorCode;
			if (!Utility.isNullOrUndefined(errorObj) && typeof (errorObj)==="object" && errorObj.error) {
				errorCode=errorObj.error.code;
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithDetails, [responseInfo.statusCode.toString(), errorObj.error.code, errorObj.error.message]);
			}
			else {
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
			}
			if (Utility.isNullOrEmptyString(errorCode)) {
				errorCode=OfficeExtension.ErrorCodes.connectionFailure;
			}
			return { errorCode: errorCode, errorMessage: errorMessage };
		};
		Utility._copyHeaders=function (src, dest) {
			if (src && dest) {
				for (var key in src) {
					dest[key]=src[key];
				}
			}
		};
		Utility._toCamelLowerCase=function (name) {
			if (Utility.isNullOrEmptyString(name)) {
				return name;
			}
			var index=0;
			while (index < name.length && name.charCodeAt(index) >=65 && name.charCodeAt(index) <=90) {
				index++;
			}
			if (index < name.length) {
				return name.substr(0, index).toLowerCase()+name.substr(index);
			}
			else {
				return name.toLowerCase();
			}
		};
		Utility.definePropertyThrowUnloadedException=function (obj, typeName, propertyName) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					throw Utility.createPropertyNotLoadedException(typeName, propertyName);
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
				}
			});
		};
		Utility.defineReadOnlyPropertyWithValue=function (obj, propertyName, value) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					return value;
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
				}
			});
		};
		Utility.processRetrieveResult=function (proxy, value, result, childItemCreateFunc) {
			if (Utility.isNullOrUndefined(value)) {
				return;
			}
			if (childItemCreateFunc) {
				var data=value[OfficeExtension.Constants.itemsLowerCase];
				if (Array.isArray(data)) {
					var itemsResult=[];
					for (var i=0; i < data.length; i++) {
						var itemProxy=childItemCreateFunc(data[i], i);
						var itemResult={};
						itemResult[OfficeExtension.Constants.proxy]=itemProxy;
						itemProxy._handleRetrieveResult(data[i], itemResult);
						itemsResult.push(itemResult);
					}
					Utility.defineReadOnlyPropertyWithValue(result, OfficeExtension.Constants.itemsLowerCase, itemsResult);
				}
			}
			else {
				var scalarPropertyNames=proxy[OfficeExtension.Constants.scalarPropertyNames];
				var navigationPropertyNames=proxy[OfficeExtension.Constants.navigationPropertyNames];
				var typeName=proxy[OfficeExtension.Constants.className];
				if (scalarPropertyNames) {
					for (var i=0; i < scalarPropertyNames.length; i++) {
						var propName=scalarPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
						}
					}
				}
				if (navigationPropertyNames) {
					for (var i=0; i < navigationPropertyNames.length; i++) {
						var propName=navigationPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							var propProxy=proxy[propName];
							var propResult={};
							propProxy._handleRetrieveResult(propValue, propResult);
							propResult[OfficeExtension.Constants.proxy]=propProxy;
							if (Array.isArray(propResult[OfficeExtension.Constants.itemsLowerCase])) {
								propResult=propResult[OfficeExtension.Constants.itemsLowerCase];
							}
							Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
						}
					}
				}
			}
		};
		Utility.fieldName_m__items="m__items";
		Utility.fieldName_isCollection="_isCollection";
		Utility._logEnabled=false;
		Utility._synchronousCleanup=false;
		Utility._doApiNotSupportedCheck=false;
		Utility.s_underscoreCharCode="_".charCodeAt(0);
		return Utility;
	}());
	OfficeExtension.Utility=Utility;
})(OfficeExtension || (OfficeExtension={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="AgaveVisualApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _loadAndSync=OfficeExtension.Utility.loadAndSync;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _retrieveAndSync=OfficeExtension.Utility.retrieveAndSync;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeBiShim="BiShim";
	var BiShim=(function (_super) {
		__extends(BiShim, _super);
		function BiShim() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BiShim.prototype, "_className", {
			get: function () {
				return "BiShim";
			},
			enumerable: true,
			configurable: true
		});
		BiShim.prototype.initialize=function (capabilities) {
			_createMethodAction(this.context, this, "Initialize", 0, [capabilities], false);
		};
		BiShim.prototype.uninitialize=function () {
			_createMethodAction(this.context, this, "Uninitialize", 0, [], false);
		};
		BiShim.prototype.getData=function () {
			var action=_createMethodAction(this.context, this, "getData", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		BiShim.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		BiShim.newObject=function (context) {
			var ret=new OfficeCore.BiShim(context, _createNewObjectObjectPath(context, "Microsoft.AgaveVisual.BiShim", false, false));
			return ret;
		};
		BiShim.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return BiShim;
	}(OfficeExtension.ClientObject));
	OfficeCore.BiShim=BiShim;
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes["generalException"]="GeneralException";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="ExperimentApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var FlightingService=(function (_super) {
		__extends(FlightingService, _super);
		function FlightingService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FlightingService.prototype, "_className", {
			get: function () {
				return "FlightingService";
			},
			enumerable: true,
			configurable: true
		});
		FlightingService.prototype.getClientSessionId=function () {
			var action=_createMethodAction(this.context, this, "GetClientSessionId", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getDeferredFlights=function () {
			var action=_createMethodAction(this.context, this, "GetDeferredFlights", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getFeature=function (featureName, type, defaultValue, possibleValues) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null));
		};
		FlightingService.prototype.getFeatureGate=function (featureName, scope) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeatureGate", 1, [featureName, scope], false, false, null));
		};
		FlightingService.prototype.resetOverride=function (featureName) {
			_createMethodAction(this.context, this, "ResetOverride", 0, [featureName]);
		};
		FlightingService.prototype.setOverride=function (featureName, type, value) {
			_createMethodAction(this.context, this, "SetOverride", 0, [featureName, type, value]);
		};
		FlightingService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		FlightingService.newObject=function (context) {
			var ret=new OfficeCore.FlightingService(context, _createNewObjectObjectPath(context, "Microsoft.Experiment.FlightingService", false));
			return ret;
		};
		FlightingService.prototype.toJSON=function () {
			return {};
		};
		return FlightingService;
	}(OfficeExtension.ClientObject));
	OfficeCore.FlightingService=FlightingService;
	var ABType=(function (_super) {
		__extends(ABType, _super);
		function ABType() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ABType.prototype, "_className", {
			get: function () {
				return "ABType";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, "ABType", this._isNull);
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ABType.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
		};
		ABType.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ABType.prototype.toJSON=function () {
			return {
				"value": this.m_value
			};
		};
		return ABType;
	}(OfficeExtension.ClientObject));
	OfficeCore.ABType=ABType;
	var FeatureType;
	(function (FeatureType) {
		FeatureType.boolean="Boolean";
		FeatureType.integer="Integer";
		FeatureType.string="String";
	})(FeatureType=OfficeCore.FeatureType || (OfficeCore.FeatureType={}));
	var ExperimentErrorCodes;
	(function (ExperimentErrorCodes) {
		ExperimentErrorCodes.generalException="GeneralException";
	})(ExperimentErrorCodes=OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var FirstPartyApis=(function () {
		function FirstPartyApis(context) {
			this.context=context;
		}
		Object.defineProperty(FirstPartyApis.prototype, "authentication", {
			get: function () {
				if (!this.m_authentication) {
					this.m_authentication=OfficeCore.AuthenticationService.newObject(this.context);
				}
				return this.m_authentication;
			},
			enumerable: true,
			configurable: true
		});
		return FirstPartyApis;
	}());
	OfficeCore.FirstPartyApis=FirstPartyApis;
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			return _super.call(this, url) || this;
		}
		Object.defineProperty(RequestContext.prototype, "firstParty", {
			get: function () {
				if (!this.m_firstPartyApis) {
					this.m_firstPartyApis=new FirstPartyApis(this);
				}
				return this.m_firstPartyApis;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flighting", {
			get: function () {
				return this.flightingService;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "telemetry", {
			get: function () {
				if (!this.m_telemetry) {
					this.m_telemetry=OfficeCore.TelemetryService.newObject(this);
				}
				return this.m_telemetry;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "bi", {
			get: function () {
				if (!this.m_biShim) {
					this.m_biShim=OfficeCore.BiShim.newObject(this);
				}
				return this.m_biShim;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flightingService", {
			get: function () {
				if (!this.m_flightingService) {
					this.m_flightingService=OfficeCore.FlightingService.newObject(this);
				}
				return this.m_flightingService;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeExtension.ClientRequestContext));
	OfficeCore.RequestContext=RequestContext;
	function run(arg1, arg2) {
		return OfficeExtension.ClientRequestContext._runBatch("OfficeCore.run", arguments, function (requestInfo) { return new OfficeCore.RequestContext(requestInfo); });
	}
	OfficeCore.run=run;
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="TelemetryApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeTelemetryService="TelemetryService";
	var TelemetryService=(function (_super) {
		__extends(TelemetryService, _super);
		function TelemetryService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TelemetryService.prototype, "_className", {
			get: function () {
				return "TelemetryService";
			},
			enumerable: true,
			configurable: true
		});
		TelemetryService.prototype.sendTelemetryEvent=function (telemetryProperties, eventName, eventContract, eventFlags, value) {
			_createMethodAction(this.context, this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], false);
		};
		TelemetryService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		TelemetryService.newObject=function (context) {
			var ret=new OfficeCore.TelemetryService(context, _createNewObjectObjectPath(context, "Microsoft.Telemetry.TelemetryService", false, false));
			return ret;
		};
		TelemetryService.prototype.toJSON=function () {
			return {};
		};
		return TelemetryService;
	}(OfficeExtension.ClientObject));
	OfficeCore.TelemetryService=TelemetryService;
	var TelemetryErrorCodes;
	(function (TelemetryErrorCodes) {
		TelemetryErrorCodes.generalException="GeneralException";
	})(TelemetryErrorCodes=OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeFirstPartyAuth;
(function (OfficeFirstPartyAuth) {
	function getAccessToken(options) {
		var context=new OfficeCore.RequestContext();
		var auth=OfficeCore.AuthenticationService.newObject(context);
		context._customData="WacPartition";
		var promise=new OfficeExtension.Promise(function (resolve, reject) {
			var result=auth.getAccessToken(options);
			context.sync()
				.then(function () {
				resolve(result);
			})
				.catch(function (e) {
				throw e;
			});
		});
		return promise.then(function (accessTokenResult) {
			return new OfficeExtension.Promise(function (resolve, reject) {
				resolve(accessTokenResult);
			});
		});
	}
	OfficeFirstPartyAuth.getAccessToken=getAccessToken;
})(OfficeFirstPartyAuth || (OfficeFirstPartyAuth={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="Office";
	var _defaultApiSetName="OfficeSharedApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var IdentityType;
	(function (IdentityType) {
		IdentityType["organizationAccount"]="OrganizationAccount";
		IdentityType["microsoftAccount"]="MicrosoftAccount";
	})(IdentityType=OfficeCore.IdentityType || (OfficeCore.IdentityType={}));
	var _typeAuthenticationService="AuthenticationService";
	var AuthenticationService=(function (_super) {
		__extends(AuthenticationService, _super);
		function AuthenticationService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(AuthenticationService.prototype, "_className", {
			get: function () {
				return "AuthenticationService";
			},
			enumerable: true,
			configurable: true
		});
		AuthenticationService.prototype.getAccessToken=function (tokenParameters) {
			var action=_createMethodAction(this.context, this, "GetAccessToken", 1, [tokenParameters], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype.getPrimaryIdentityInfo=function () {
			_throwIfApiNotSupported("AuthenticationService.getPrimaryIdentityInfo", "FirstPartyAuthentication", "1.2", _hostName);
			var action=_createMethodAction(this.context, this, "GetPrimaryIdentityInfo", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		AuthenticationService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		AuthenticationService.newObject=function (context) {
			var ret=new OfficeCore.AuthenticationService(context, _createNewObjectObjectPath(context, "Microsoft.Authentication.AuthenticationService", false, false));
			return ret;
		};
		AuthenticationService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return AuthenticationService;
	}(OfficeExtension.ClientObject));
	OfficeCore.AuthenticationService=AuthenticationService;
	var _typeComment="Comment";
	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Comment.prototype, "_className", {
			get: function () {
				return "Comment";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "text", "created", "level", "resolved", "author", "mentions"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, false, false, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["parent", "parentOrNullObject", "replies"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parent", {
			get: function () {
				if (!this._P) {
					this._P=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "Parent", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parentOrNullObject", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "ParentOrNullObject", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "replies", {
			get: function () {
				if (!this._R) {
					this._R=new OfficeCore.CommentCollection(this.context, _createPropertyObjectPath(this.context, this, "Replies", true, false, false));
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "author", {
			get: function () {
				_throwIfNotLoaded("author", this._A, _typeComment, this._isNull);
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "created", {
			get: function () {
				_throwIfNotLoaded("created", this._C, _typeComment, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeComment, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "level", {
			get: function () {
				_throwIfNotLoaded("level", this._L, _typeComment, this._isNull);
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "mentions", {
			get: function () {
				_throwIfNotLoaded("mentions", this._M, _typeComment, this._isNull);
				return this._M;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "resolved", {
			get: function () {
				_throwIfNotLoaded("resolved", this._Re, _typeComment, this._isNull);
				return this._Re;
			},
			set: function (value) {
				this._Re=value;
				_createSetPropertyAction(this.context, this, "Resolved", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeComment, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["text", "resolved"], [], [
				"parent",
				"parentOrNullObject",
				"replies"
			]);
		};
		Comment.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Comment.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Comment.prototype.getParentOrSelf=function () {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "GetParentOrSelf", 1, [], false, false, null, false));
		};
		Comment.prototype.getRichText=function (format) {
			var action=_createMethodAction(this.context, this, "GetRichText", 1, [format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype.reply=function (text, format) {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "Reply", 0, [text, format], false, false, null, false));
		};
		Comment.prototype.setRichText=function (text, format) {
			var action=_createMethodAction(this.context, this, "SetRichText", 0, [text, format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Author"])) {
				this._A=obj["Author"];
			}
			if (!_isUndefined(obj["Created"])) {
				this._C=_adjustToDateTime(obj["Created"]);
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Level"])) {
				this._L=obj["Level"];
			}
			if (!_isUndefined(obj["Mentions"])) {
				this._M=obj["Mentions"];
			}
			if (!_isUndefined(obj["Resolved"])) {
				this._Re=obj["Resolved"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			_handleNavigationPropertyResults(this, obj, ["parent", "Parent", "parentOrNullObject", "ParentOrNullObject", "replies", "Replies"]);
		};
		Comment.prototype.load=function (option) {
			return _load(this, option);
		};
		Comment.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Comment.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Comment.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			if (!_isUndefined(obj["Created"])) {
				obj["created"]=_adjustToDateTime(obj["created"]);
			}
			_processRetrieveResult(this, value, result);
		};
		Comment.prototype.toJSON=function () {
			return _toJson(this, {
				"author": this._A,
				"created": this._C,
				"id": this._I,
				"level": this._L,
				"mentions": this._M,
				"resolved": this._Re,
				"text": this._T,
			}, {
				"replies": this._R,
			});
		};
		Comment.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Comment;
	}(OfficeExtension.ClientObject));
	OfficeCore.Comment=Comment;
	var _typeCommentCollection="CommentCollection";
	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentCollection.prototype, "_className", {
			get: function () {
				return "CommentCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCommentCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		CommentCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CommentCollection.prototype.getItem=function (id) {
			return new OfficeCore.Comment(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		CommentCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new OfficeCore.Comment(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CommentCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CommentCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CommentCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new OfficeCore.Comment(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		CommentCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return CommentCollection;
	}(OfficeExtension.ClientObject));
	OfficeCore.CommentCollection=CommentCollection;
	var CommentTextFormat;
	(function (CommentTextFormat) {
		CommentTextFormat["plain"]="Plain";
		CommentTextFormat["markdown"]="Markdown";
		CommentTextFormat["delta"]="Delta";
	})(CommentTextFormat=OfficeCore.CommentTextFormat || (OfficeCore.CommentTextFormat={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes["apiNotAvailable"]="ApiNotAvailable";
		ErrorCodes["clientError"]="ClientError";
		ErrorCodes["invalidArgument"]="InvalidArgument";
		ErrorCodes["invalidGrant"]="InvalidGrant";
		ErrorCodes["invalidResourceUrl"]="InvalidResourceUrl";
		ErrorCodes["serverError"]="ServerError";
		ErrorCodes["unsupportedUserIdentity"]="UnsupportedUserIdentity";
		ErrorCodes["userNotSignedIn"]="UserNotSignedIn";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
