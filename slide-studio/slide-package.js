/* Build one slide from the toolkit's PowerPoint-authored content-add-in slide. */
(() => {
  'use strict';
  const ns = 'http://schemas.microsoft.com/office/webextensions/webextension/2010/11';
  const extPath = 'ppt/webextensions/webextension1.xml';
  const slidePath = 'ppt/slides/slide1.xml';
  const relPath = 'ppt/slides/_rels/slide1.xml.rels';
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  let template;
  function base64(bytes) {
    let text = '';
    for (let i=0; i<bytes.length; i+=16384) text += String.fromCharCode(...bytes.subarray(i,i+16384));
    return btoa(text);
  }
  function parse(text) {
    const xml = new DOMParser().parseFromString(text,'application/xml');
    if (xml.getElementsByTagName('parsererror').length) throw Error('The XML could not be read.');
    return xml;
  }
  function nodes(xml, name) { return Array.from(xml.getElementsByTagNameNS('*',name)); }
  function snapshot(title) {
    const canvas = document.createElement('canvas');
    canvas.width=960; canvas.height=540;
    const g=canvas.getContext('2d');
    g.fillStyle='#f4f6f9';g.fillRect(0,0,960,540);
    g.fillStyle='#911337';g.fillRect(0,0,14,540);
    g.fillStyle='#1c3048';g.font='600 26px sans-serif';g.fillText('QUB Slide Studio',48,72);
    g.font='600 42px sans-serif';g.fillText(title,48,245,850);
    g.font='24px sans-serif';g.fillText('Start Slide Show with Slide Web Viewer available.',48,304,850);
    return Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]),c=>c.charCodeAt(0));
  }
  async function build(activity, catalog, viewer) {
    if (!template) {
      const response = await fetch('activity-template.pptx');
      if (!response.ok) throw Error('Could not load the slide template. Check your connection and try again.');
      template = new Uint8Array(await response.arrayBuffer());
    }
    const files = fflate.unzipSync(template);
    const ext = parse(dec.decode(files[extPath]));
    const slide = parse(dec.decode(files[slidePath]));
    const rel = parse(dec.decode(files[relPath]));
    ext.documentElement.setAttribute('id','{'+crypto.randomUUID().toUpperCase()+'}');
    const reference=nodes(ext,'reference')[0];
    for (const [key,value] of Object.entries({id:viewer.id,version:viewer.version,store:catalog.store,storeType:catalog.type})) reference.setAttribute(key,value);
    const property=ext.createElementNS(ns,'we:property');
    property.setAttribute('name',activity.html?'qubSlideStudio.html.v1':'slideWebViewer.url.v1');
    property.setAttribute('value',JSON.stringify(activity.html?base64(enc.encode(activity.html)):activity.url));
    nodes(ext,'properties')[0].append(property);
    nodes(slide,'cSld')[0].setAttribute('name',activity.title);
    nodes(slide,'creationId')[0].setAttribute('val',String(crypto.getRandomValues(new Uint32Array(1))[0]));
    for (const element of nodes(slide,'cNvPr').filter(el=>el.getAttribute('id')==='4')) {
      element.setAttribute('name',activity.title);
      element.setAttribute('descr','Interactive teaching activity: '+activity.title);
    }
    const text = activity.title + (activity.url?' · Open in browser':'');
    nodes(slide,'t')[0].textContent=text;
    const link=nodes(rel,'Relationship').find(item=>item.getAttribute('Type').endsWith('/hyperlink'));
    if (activity.url) link.setAttribute('Target',activity.url);
    else {
      link.remove();
      nodes(slide,'hlinkClick').forEach(item=>item.remove());
      nodes(slide,'rPr').forEach(item=>item.removeAttribute('u'));
    }
    const serializer=new XMLSerializer();
    for (const [name,xml] of [[extPath,ext],[slidePath,slide],[relPath,rel]]) files[name]=enc.encode(serializer.serializeToString(xml));
    files['ppt/media/image1.png']=snapshot(activity.title);
    return fflate.zipSync(files,{level:4});
  }
  globalThis.StudioPackage=Object.freeze({build,base64,parse,nodes});
})();
