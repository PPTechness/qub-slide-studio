'use strict';
const el = id => document.getElementById(id);
const {pageAddress,addressFromEntry} = SlideAddresses;
const {parse,nodes} = StudioPackage;
const standard = {id:'ba0c1557-8450-4520-91d8-c5428baa5cd7',version:'1.2.0.0',name:'Slide Web Viewer'};
const officeNS = 'http://schemas.microsoft.com/office/appforoffice/1.1';
const library = StudioLibrary;
const limit = 1024*1024;
const launch = new URLSearchParams(location.search);
const previewOnly = launch.get('preview') === '1';
const managed = launch.get('deployment') === 'managed';
let viewer = standard, platform='', active=false, busy=false, revision=0;
const options = (select,items) => select.replaceChildren(...items.map(([value,label])=>new Option(label,value)));
const problem = message => {el('error').textContent=message;};
function preferences() {
  try {return JSON.parse(localStorage.getItem('qub.slideStudio.setup.v1')) || {};} catch {return {};}
}
function remember() {
  try {localStorage.setItem('qub.slideStudio.setup.v1',JSON.stringify({catalog:el('catalog').value,share:el('share').value.trim(),viewer}));} catch {}
}
function showSetup() {
  el('share-field').hidden=el('catalog').value!=='share';
  el('using').textContent='On-slide viewer: '+viewer.name;
  el('reset-player').hidden=viewer.id===standard.id;
}
const saved=preferences();
if (!managed && saved.viewer && saved.viewer.id!==standard.id && /^[\da-f-]{36}$/i.test(saved.viewer.id) && /^\d+(\.\d+){3}$/.test(saved.viewer.version)) viewer=saved.viewer;
el('share').value=saved.share || '';
function catalogRef() {
  const catalog=managed?'managed':el('catalog').value;
  if (catalog==='managed') return {type:'EXCatalog',store:'EXCatalog'};
  if (catalog==='mac') return {type:'Registry',store:'developer'};
  let path=el('share').value.trim();
  if (/^file:\/\//i.test(path)) {
    try {
      const directory=new URL(path);
      if (directory.username || directory.password || directory.search || directory.hash) throw Error();
      path='\\\\'+directory.hostname+decodeURIComponent(directory.pathname).replaceAll('/','\\');
    } catch {path='';}
  }
  path=path.replace(/\\$/,'');
  const folders=path.slice(2).split('\\');
  if (!path.startsWith('\\\\') || folders.length<2 || !/^[a-z\d._-]+$/i.test(folders[0]) ||
      folders.some(part=>!part || part==='.' || part==='..' || /[/:*?"<>|\x00-\x1f]/.test(part))) {
    el('setup').open=true;
    throw Error('Under Viewer setup, paste the network address of your WorkshopAddins folder, then insert again.');
  }
  return {type:'FileSystem',store:path};
}
function readManifest(text) {
  if (/<!DOCTYPE/i.test(text)) throw Error('Choose a PowerPoint content add-in XML file.');
  const xml=parse(text);
  const root=xml.documentElement;
  if (root.namespaceURI!==officeNS || root.localName!=='OfficeApp' ||
      root.getAttributeNS('http://www.w3.org/2001/XMLSchema-instance','type')?.split(':').pop()!=='ContentApp' ||
      !nodes(xml,'Host').some(h=>h.getAttribute('Name')==='Presentation')) throw Error('Choose a fixed-page PowerPoint content add-in XML.');
  const source=xml.getElementsByTagNameNS(officeNS,'DefaultSettings')[0]?.getElementsByTagNameNS(officeNS,'SourceLocation')[0];
  const url=pageAddress(source?.getAttribute('DefaultValue'));
  if (!url) throw Error('This manifest needs a shareable HTTPS starting address.');
  return {url, id:nodes(xml,'Id')[0]?.textContent.trim(),version:nodes(xml,'Version')[0]?.textContent.trim(),name:nodes(xml,'DisplayName')[0]?.getAttribute('DefaultValue') || 'Web activity'};
}
el('player').onchange=async event=>{
  try {
    const file=event.target.files[0];if (!file) return;
    if (file.size>65536) throw Error('Choose a Slide Web Viewer XML under 64 KB.');
    const v=readManifest(await file.text());
    if (!/^slide web viewer( \(personal\))?$/i.test(v.name) || !/\/slide-web-viewer\/viewer(?:\.html)?$/.test(new URL(v.url).pathname) ||
        !/^[\da-f]{8}(-[\da-f]{4}){3}-[\da-f]{12}$/i.test(v.id) || !/^\d{1,5}(\.\d{1,5}){3}$/.test(v.version)) throw Error('Choose the standard or personal Slide Web Viewer XML you installed.');
    viewer={id:v.id,version:v.version,name:v.name};remember();showSetup();problem('');
  } catch(e){problem(e.message);} finally {event.target.value='';}
};
el('reset-player').onclick=()=>{viewer=standard;remember();showSetup();};
el('catalog').onchange=()=>{showSetup();remember();};el('share').onchange=remember;

function invalidate() {
  revision++;problem('');
  if (active) el('status').textContent='';
  el('preview').open=false;el('preview-area').replaceChildren();
}
function configureEdition() {
  const kind=el('tool').value,own=el('edition').value==='own';
  el('own').hidden=!own;
  el('preset').hidden=own && kind==='bingo';
  el('variant-label').textContent=kind==='bingo'?'Starting round':'Starting view';
  options(el('variant'),kind==='bingo'?(library.bingo[el('edition').value]?.rounds||[]).map(r=>[r,r]):[['recall','Recall from a clue'],['read','Read notes']]);
  el('entries-label').textContent=kind==='bingo'?'Terms and clues':'Sticky notes';
  el('entry-hint').textContent=kind==='bingo'?'Enter at least four different terms, one per line: term | clue.':'Enter one note per line: term | explanation | recall clue. The clue is optional.';
  el('entries').placeholder=kind==='bingo'?'Archive | Where past records are kept':'Archive | A collection of records | Where are past records kept?';
  invalidate();
}
const drafts={};let previousKind='web';
function configureTool() {
  const kind=el('tool').value;
  const shortcut=el('tool').selectedOptions[0]?.dataset.page;
  el('shortcut').hidden=!shortcut;
  el('shortcut').textContent=shortcut ? 'Opens '+new URL(shortcut).hostname+' on your slide. Use Webpage or embed for a specific link.' : '';
  drafts[previousKind]=el('entries').value;el('entries').value=drafts[kind]||'';previousKind=kind;
  el('web').hidden=kind!=='web';el('upload').hidden=kind!=='file';el('activity').hidden=!library[kind];
  if (library[kind]) {
    options(el('edition'),[...Object.entries(library[kind]).map(([id,row])=>[id,row.label]),['own',kind==='bingo'?'Write my own terms':'Write my own notes']]);
    el('edition').value='queens';configureEdition();
  }
  invalidate();
}
el('tool').onchange=configureTool;el('edition').onchange=configureEdition;
for(const id of ['address','entries','title','variant','file']) el(id).addEventListener('input',invalidate);
function rows(kind) {
  const seen=new Set();
  const result=el('entries').value.split(/\r?\n/).filter(line=>line.trim()).map((line,i)=>{
    const cells=line.split('|').map(s=>s.trim());
    if (!cells[0]||!cells[1]||cells.length>(kind==='bingo'?2:3)) throw Error('Check line '+(i+1)+'. Follow the format shown above the box.');
    if (seen.has(cells[0].toLowerCase())) throw Error('Use a different term on line '+(i+1)+'.');
    seen.add(cells[0].toLowerCase());return cells;
  });
  if (result.length<(kind==='bingo'?4:1)) throw Error(kind==='bingo'?'Add at least four terms and clues for a bingo card.':'Add at least one sticky note.');
  return result;
}
const slidesBase='../vibe-coding/slides/';
// A deck downloaded from a local preview still needs a portable teaching-page link.
const publishedRoot='https://pptechness.github.io/qub-slide-studio/slide-studio/';
async function customActivity(kind,title) {
  const items=rows(kind);
  const response=await fetch(slidesBase+kind+'.html');
  if(!response.ok) throw Error('Could not load the teaching activity. Check your connection and try again.');
  let markup=await response.text();
  const id=kind==='bingo'?'bank-data':'book-data';
  const tag=new RegExp('(<script[^>]*\\bid="'+id+'"[^>]*>)[\\s\\S]*?(</script>)');
  if(!tag.test(markup)) throw Error('The activity template is unavailable.');
  const data=kind==='bingo'?{custom:{title,credit:'',terms:items.map(([term,clue])=>({term,clue,definition:clue})),rounds:{'My terms':items.map(([term])=>term)}}}:
    {start:'custom',custom:true,mode:'learner',study:el('variant').value,boards:[{id:'custom',tab:title,name:title,question:'Explore the notes or switch to recall.',credit:'',topics:['My notes'],notes:items.map(([term,body,clue],i)=>({topic:0,colour:['Honey','Mint','Lilac','Rose'][i%4],body:term+'\n'+body,clue:clue||body,url:'',linkText:''}))}]};
  markup=markup.replace(tag,(_,open,close)=>open+JSON.stringify(data).replace(/</g,'\\u003c')+close);
  const fullLink=kind==='bingo'?'open-full':'full-screen';
  markup=markup.replace('id="'+fullLink+'"','id="'+fullLink+'" hidden');
  return markup;
}
function capped(html) {
  if(!html.trim() || new TextEncoder().encode(html).length>limit) throw Error('Choose a self-contained UTF-8 HTML file up to 1 MiB, or paste its hosted URL.');
  return html;
}
async function activity() {
  const kind=el('tool').value,title=el('title').value.trim();
  const option=el('tool').selectedOptions[0];
  if (option.dataset.page) return {url:option.dataset.page,title:title||option.textContent};
  if (kind==='web') {
    const url=addressFromEntry(el('address').value);
    if (!url) throw Error('Paste an HTTPS page address or a single iframe with an HTTPS source.');
    return {url,title:title||new URL(url).hostname};
  }
  if (kind==='file') {
    const file=el('file').files[0];
    if(!file) throw Error('Choose an HTML activity or a fixed-page XML.');
    if(file.size>limit) throw Error('For a file over 1 MiB, paste the hosted page address.');
    const text=new TextDecoder('utf-8',{fatal:true}).decode(await file.arrayBuffer());
    if(/\.html?$/i.test(file.name)) return {html:capped(text),title:title||file.name.replace(/\.html?$/i,'')};
    if(!/\.xml$/i.test(file.name)) throw Error('Choose an .html, .htm or .xml file.');
    const v=readManifest(text);
    if(/\/slide-web-viewer\/viewer(?:\.html)?$/.test(new URL(v.url).pathname)) throw Error('This XML installs the reusable viewer. Paste the teaching page itself, or choose a fixed-page XML.');
    return {url:v.url,title:title||v.name};
  }
  const edition=el('edition').value;
  if(edition==='own') {
    const name=title||(kind==='bingo'?'My vocabulary bingo':'My sticky notes');
    return {html:capped(await customActivity(kind,name)),title:name};
  }
  const base=location.protocol==='https:'?location.href:publishedRoot;
  const url=new URL(slidesBase+kind+'.html',base);
  url.hash=edition;
  url.searchParams.set(kind==='bingo'?'round':'study',el('variant').value);
  if(kind==='qblit') url.searchParams.set('learner','1');
  return {url:url.href,title:title||library[kind][edition].label};
}
el('preview').addEventListener('toggle',async()=>{
  const area=el('preview-area');
  area.replaceChildren();
  if(!el('preview').open) return;
  const token=revision;
  try{
    const a=await activity();if(token!==revision||!el('preview').open) return;
    const frame=document.createElement('iframe');frame.title='Preview of '+a.title;
    frame.setAttribute('sandbox','allow-scripts allow-forms');frame.referrerPolicy='no-referrer';
    if(a.html)frame.srcdoc=a.html;else frame.src=a.url;
    area.append(frame);
  }catch(e){area.textContent=e.message;}
});
function selectedSlide() {
  return new Promise((resolve,reject)=>Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,result=>{
    const id=result.value?.slides?.[0]?.id;
    if(result.status!==Office.AsyncResultStatus.Succeeded||id==null) reject(Error('Select a slide thumbnail, then choose Insert new slide.'));
    else resolve(String(id).split('#')[0]+'#');
  }));
}
function download(bytes,title) {
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}));
  a.download=(title.replace(/[^a-z0-9-]+/gi,' ').trim().slice(0,70)||'Teaching activity')+'.pptx';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),30000);
}
el('studio').addEventListener('submit',async event=>{
  event.preventDefault();if(busy||!active) return;
  busy=true;el('insert').disabled=true;el('controls').disabled=true;problem('');el('status').textContent='Preparing your slide…';
  try {
    const place=catalogRef(),chosen=viewer;
    const a=await activity();
    const bytes=await StudioPackage.build(a,place,chosen);
    if(previewOnly) {download(bytes,a.title);el('status').textContent='Your test slide has been downloaded.';}
    else {
      const target=await selectedSlide();
      await PowerPoint.run(async context=>{
        context.presentation.insertSlidesFromBase64(StudioPackage.base64(bytes),{targetSlideId:target,formatting:'KeepSourceFormatting'});
        await context.sync();
      });
      el('status').textContent='Inserted “'+a.title+'” after the selected slide. Start Slide Show to try it.';
    }
  }catch(e){el('status').textContent='';problem(e.message||'PowerPoint could not insert the slide. Check the presentation is editable.');}
  finally{busy=false;el('insert').disabled=false;el('controls').disabled=false;}
});
function ready() {active=true;el('insert').disabled=false;showSetup();}
el('catalog').value=saved.catalog||'mac';showSetup();configureTool();
if(previewOnly) {
  ready();el('insert').textContent='Download test slide';el('status').textContent='Browser preview. Download a slide to try in PowerPoint.';
} else {
  const timer=setTimeout(()=>{if(!active) el('status').textContent='Open this panel from PowerPoint, or use the browser preview in the pilot guide.';},12000);
  window.Office?.onReady().then(info=>{
    clearTimeout(timer);
    if(info.host!==Office.HostType.PowerPoint || !Office.context.requirements.isSetSupported('PowerPointApi','1.2')) {
      el('status').textContent='Open QUB Slide Studio in a version of PowerPoint that supports slide insertion.';return;
    }
    platform=info.platform;
    if(managed) { el('catalog').value='managed'; el('setup').hidden=true; }
    else if(!saved.catalog) el('catalog').value=platform===Office.PlatformType.Mac?'mac':platform===Office.PlatformType.PC?'share':'managed';
    ready();el('status').textContent='';
  }).catch(()=>{el('status').textContent='PowerPoint could not start the panel. Reopen the presentation and try again.';});
}
