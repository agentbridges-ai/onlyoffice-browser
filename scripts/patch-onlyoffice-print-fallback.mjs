#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const roots = process.argv.slice(2);
if (roots.length === 0) roots.push('dist');

const printDocumentMain = `this.iframePrint.onload=function(){var t=e.iframePrint;try{t.style.display="block",t.style.visibility="visible",t.style.position="fixed",t.style.left="0",t.style.top="0",t.style.width="100vw",t.style.height="100vh",t.style.opacity="0.01",t.style.pointerEvents="none",t.style.border="0",t.style.zIndex="-1",setTimeout((function(){var e=document.title;try{try{document.title=decodeURIComponent((new URL(t.src,location.href).pathname.split("/").pop()||"").replace(/\\.pdf$/i,""))||e}catch(t){}t.contentWindow.focus(),t.contentWindow.print(),t.contentWindow.blur(),window.focus()}catch(t){try{console.warn("OnlyOffice browser PDF print failed",t)}catch(e){}}finally{document.title=e}}),300)}catch(t){try{console.warn("OnlyOffice browser PDF print failed",t)}catch(e){}}}`;
const printEmbed = `o.onload=function(){var e=o;try{e.style.display="block",e.style.visibility="visible",e.style.position="fixed",e.style.left="0",e.style.top="0",e.style.width="100vw",e.style.height="100vh",e.style.opacity="0.01",e.style.pointerEvents="none",e.style.border="0",e.style.zIndex="-1",setTimeout((function(){var t=document.title;try{try{document.title=decodeURIComponent((new URL(e.src,location.href).pathname.split("/").pop()||"").replace(/\\.pdf$/i,""))||t}catch(e){}e.contentWindow.focus(),e.contentWindow.print(),e.contentWindow.blur(),window.focus()}catch(t){try{console.warn("OnlyOffice browser PDF print failed",t)}catch(e){}}finally{document.title=t}}),300)}catch(t){try{console.warn("OnlyOffice browser PDF print failed",t)}catch(e){}}}`;

const replacements = [
  {
    file: 'web-apps/apps/documenteditor/main/app.js',
    from: 'this.iframePrint.onload=function(){try{e.iframePrint.contentWindow.focus(),e.iframePrint.contentWindow.print(),e.iframePrint.contentWindow.blur(),window.focus()}catch(e){window.open(t,"_blank")}}',
    to: printDocumentMain,
  },
  {
    file: 'web-apps/apps/presentationeditor/main/app.js',
    from: 'this.iframePrint.onload=function(){try{e.iframePrint.contentWindow.focus(),e.iframePrint.contentWindow.print(),e.iframePrint.contentWindow.blur(),window.focus()}catch(e){window.open(t,"_blank")}}',
    to: printDocumentMain,
  },
  {
    file: 'web-apps/apps/spreadsheeteditor/main/app.js',
    from: 'this.iframePrint.onload=function(){try{e.iframePrint.contentWindow.focus(),e.iframePrint.contentWindow.print(),e.iframePrint.contentWindow.blur(),window.focus()}catch(i){var t=new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.PDF);t.asc_setAdvancedOptions(e.getApplication().getController("Print").getPrintParams()),e.api.asc_DownloadAs(t)}}',
    to: printDocumentMain,
  },
  {
    file: 'web-apps/apps/documenteditor/embed/app-all.js',
    from: 'o.onload=function(){try{o.contentWindow.focus(),o.contentWindow.print(),o.contentWindow.blur(),window.focus()}catch(t){e.asc_DownloadAs(new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.PDF))}}',
    to: printEmbed,
  },
  {
    file: 'web-apps/apps/presentationeditor/embed/app-all.js',
    from: 'o.onload=function(){try{o.contentWindow.focus(),o.contentWindow.print(),o.contentWindow.blur(),window.focus()}catch(t){e.asc_DownloadAs(new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.PDF))}}',
    to: printEmbed,
  },
  {
    file: 'web-apps/apps/spreadsheeteditor/embed/app-all.js',
    from: 'o.onload=function(){try{o.contentWindow.focus(),o.contentWindow.print(),o.contentWindow.blur(),window.focus()}catch(e){t.asc_DownloadAs(new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.PDF))}}',
    to: printEmbed,
  },
];

let patched = 0;
for (const root of roots) {
  for (const replacement of replacements) {
    const target = path.join(root, replacement.file);
    if (!fs.existsSync(target)) continue;

    let source = fs.readFileSync(target, 'utf8');
    if (!source.includes(replacement.from)) {
      throw new Error(`OnlyOffice print fallback patch target not found: ${target}`);
    }

    source = source.replace(replacement.from, replacement.to);
    fs.writeFileSync(target, source);
    patched += 1;
  }
}

console.log(`Patched ${patched} OnlyOffice print fallback handlers.`);
