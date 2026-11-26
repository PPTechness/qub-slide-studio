'use strict';
const ns = 'http://schemas.microsoft.com/office/appforoffice/1.1';
const byId = id => document.getElementById(id);
const personalName = 'Slide Web Viewer (personal)';
const sites = new Map();
let template, viewerOrigin, personalId, lastVersion;

function report(message, failed = false) {
  byId('feedback').textContent = failed ? '' : message;
  byId('problem').textContent = failed ? message : '';
}

function parseXml(text) {
  if (text.length > 65536 || /<!DOCTYPE/i.test(text)) throw Error('Choose a viewer XML file under 64 KB.');
  const xml = new DOMParser().parseFromString(text, 'application/xml');
  if (xml.documentElement.namespaceURI !== ns || xml.documentElement.localName !== 'OfficeApp') {
    throw Error('Choose a PowerPoint viewer XML file.');
  }
  return xml;
}

function node(xml, name) {
  const matches = xml.getElementsByTagNameNS(ns, name);
  if (matches.length !== 1) throw Error('The XML is missing a valid ' + name + ' setting.');
  return matches[0];
}

function originFor(value) {
  try {
    if (!/^https:\/\//i.test(value.trim())) throw Error();
    const url = new URL(value.trim());
    if (url.protocol !== 'https:' || url.username || url.password || !/^[a-z0-9.-]+$/.test(url.hostname)) throw Error();
    return url.origin;
  } catch { throw Error('Enter an HTTPS website address without credentials or wildcards.'); }
}

function listed(xml) {
  return Array.from(xml.getElementsByTagNameNS(ns, 'AppDomain'), item => originFor(item.textContent));
}

function renderSites() {
  const rows = document.createDocumentFragment();
  for (const [site, selected] of sites) {
    const label = document.createElement('label');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = selected;
    checkbox.addEventListener('change', () => {
      sites.set(site, checkbox.checked);
      report('Download the XML to keep this selection.');
    });
    label.append(checkbox, site);
    rows.append(label);
  }
  byId('sites').replaceChildren(rows);
}

function setSites(values) {
  sites.clear();
  for (const site of values) if (site !== viewerOrigin) sites.set(site, true);
  renderSites();
}

function nextVersion() {
  if (!lastVersion) return node(template, 'Version').textContent;
  const parts = lastVersion.split('.').map(Number);
  for (let i = parts.length - 1; i >= 0; i--) {
    if (parts[i] < 65535) { parts[i]++; return parts.join('.'); }
    parts[i] = 0;
  }
  throw Error('This add-in version cannot be increased.');
}

byId('add-site').addEventListener('submit', event => {
  event.preventDefault();
  try {
    const site = originFor(byId('new-site').value);
    if (site === viewerOrigin) { report('The viewer’s own website is already included.'); return; }
    sites.set(site, true);
    renderSites();
    byId('new-site').value = '';
    report('Included ' + site + '. Download the XML to keep your selection.');
  } catch (error) { report(error.message, true); }
});

byId('import').addEventListener('change', async event => {
  const input = event.target;
  const file = input.files[0];
  if (!file) return;
  input.disabled = true;
  byId('download').disabled = true;
  try {
    if (file.size > 65536) throw Error('Choose a personal viewer XML under 64 KB.');
    const xml = parseXml(await file.text());
    const id = node(xml, 'Id').textContent.toLowerCase();
    const version = node(xml, 'Version').textContent;
    const source = new URL(node(xml, 'SourceLocation').getAttribute('DefaultValue'));
    if (node(xml, 'DisplayName').getAttribute('DefaultValue') !== personalName ||
        source.searchParams.get('personal') !== '1' || id === node(template, 'Id').textContent.toLowerCase() ||
        !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(id) ||
        !/^\d{1,5}(\.\d{1,5}){3}$/.test(version) || version.split('.').some(part => Number(part) > 65535)) {
      throw Error('Choose an XML previously downloaded from the personal viewer builder.');
    }
    const imported = listed(xml);
    setSites(imported);
    personalId = id;
    lastVersion = version;
    report('Opened your personal viewer, version ' + version + '. Your next download will update the same add-in.');
  } catch (error) { report(error.message, true); }
  finally { input.value = ''; input.disabled = false; byId('download').disabled = false; }
});

byId('download').addEventListener('click', () => {
  try {
    if (byId('new-site').value.trim()) throw Error('Choose Add website first, or clear the address before downloading.');
    const xml = template.cloneNode(true);
    const version = nextVersion();
    node(xml, 'Id').textContent = personalId;
    node(xml, 'Version').textContent = version;
    node(xml, 'DisplayName').setAttribute('DefaultValue', personalName);
    const location = node(xml, 'SourceLocation');
    const url = new URL(location.getAttribute('DefaultValue'));
    url.searchParams.delete('requestForm');
    url.searchParams.set('personal', '1');
    url.searchParams.set('direct', '1');
    url.searchParams.delete('site');
    xml.getElementsByTagNameNS(ns, 'AppDomains')[0]?.remove();
    const domains = xml.createElementNS(ns, 'AppDomains');
    for (const [site, selected] of sites) {
      if (!selected) continue;
      const entry = xml.createElementNS(ns, 'AppDomain');
      entry.textContent = site;
      domains.append(entry);
      url.searchParams.append('site', site);
    }
    if (url.href.length > 2018) throw Error('The selected websites exceed Office’s launch URL limit. Deselect some and try again.');
    if (domains.childElementCount) node(xml, 'Hosts').before(domains);
    location.setAttribute('DefaultValue', url.href);
    const blob = new Blob([new XMLSerializer().serializeToString(xml)], {type:'application/xml'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'slide-web-viewer-personal.xml';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
    lastVersion = version;
    report('Downloaded version ' + version + '. Keep this XML to update the same slides next time.');
  } catch (error) { report(error.message, true); }
});

(async () => {
  try {
    const response = await fetch('slide-web-viewer.xml', {cache:'no-cache'});
    if (!response.ok) throw Error();
    template = parseXml(await response.text());
    viewerOrigin = originFor(node(template, 'SourceLocation').getAttribute('DefaultValue'));
    personalId = crypto.randomUUID();
    setSites(listed(template));
    const guide = node(template, 'SupportUrl').getAttribute('DefaultValue');
    if (guide.startsWith('https://')) byId('guide').href = guide;
    for (const id of ['import','new-site','add','download']) byId(id).disabled = false;
    report('');
  } catch { report('Could not load the viewer template. Reopen this page over HTTPS and try again.', true); }
})();
