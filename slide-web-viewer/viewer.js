'use strict';
(async function startViewer() {
if (!globalThis.SlideAddresses) await import('./addresses.js?v=1');

const element = id => document.getElementById(id);
const frame = element('page');
const editor = element('editor');
const browserCard = element('browser-card');
const input = element('address');
const save = element('save');
const change = element('change');
const cancel = element('cancel');
const openPage = element('open-page');
const status = element('status');
const error = element('error');
const form = editor.querySelector('form');
const settingKey = 'slideWebViewer.url.v1';
const launch = new URLSearchParams(location.search);
const preview = launch.get('preview') === '1';
// The installed manifest supplies these origins. Office enforces its own AppDomains.
// A pasted page address never adds a trusted site.
const directOrigins = new Set([location.origin, ...launch.getAll('site')]);
let desktopDirect = false;
let presenting = false;
let settings;
let savedUrl = '';
let editable = false;
let editing = false;
let saving = false;
let tryFrame = false;

const {pageAddress, addressFromEntry} = globalThis.SlideAddresses;
let localHtml = '';

// Only the specific landing pages checked on 26 Nov 2026. A site may offer
// separate embed URLs. Keep all other paths eligible and allow a manual retry.
const browserFirstPages = new Map([
  ['https://www.qub.ac.uk', 'Queen’s University Belfast'],
  ['https://qub.ac.uk', 'Queen’s University Belfast'],
  ['https://www.bbc.co.uk/sport', 'BBC Sport'],
  ['https://www.bbc.co.uk/news', 'BBC News'],
  ['https://google.com', 'Google'],
  ['https://www.google.com', 'Google']
]);

function browserPageName(address) {
  if (!address) return '';
  const url = new URL(address);
  return browserFirstPages.get(url.origin + url.pathname.replace(/\/$/, '')) || '';
}

function offersDirect(address) {
  return Boolean(desktopDirect && address && directOrigins.has(new URL(address).origin));
}

function updateHint() {
  const hint = element('route-hint');
  if (!hint) return;
  const url = addressFromEntry(input.value);
  let message = '';
  if (offersDirect(url)) {
    message = 'In desktop Slide Show, choose Open on slide to open this site directly.';
  } else if (browserPageName(url)) {
    message = 'This homepage may not load in the embedded view. Open in browser is available on your slide.';
  } else if (url) {
    message = 'This page will use the embedded view. If the site blocks embedding, choose Open in browser.';
  }
  if (url && input.value.trimStart().startsWith('<')) message = 'Embed address found on ' + new URL(url).host + '. ' + message;
  if (hint.textContent !== message) hint.textContent = message;
  const request = element('request');
  if (request) {
    const visible = desktopDirect && launch.get('personal') !== '1' && url && !offersDirect(url);
    request.hidden = !visible;
    if (!visible) request.open = false;
    element('request-site').value = visible ? new URL(url).origin : '';
    const formUrl = pageAddress(launch.get('requestForm'));
    const formLink = element('request-form');
    formLink.hidden = !formUrl;
    if (formUrl) formLink.href = formUrl;
  }
}

function directLink(id, visible) {
  const link = element(id);
  if (!link) return;
  link.hidden = !visible;
  if (visible) link.href = savedUrl;
  else link.removeAttribute('href');
}

const viewport = element('viewport');
const scale = element('scale');
function sizeActivity() {
  const factor = presenting || preview ? Number(scale.value) : 1;
  viewport.style.setProperty('--activity-scale', factor);
  viewport.style.setProperty('--activity-size', (100 / factor) + '%');
  element('scale-control').hidden = viewport.hidden || !(presenting || preview);
}
scale.addEventListener('change', sizeActivity);

function render() {
  viewport.hidden = !localHtml && (!savedUrl || (editable && editing) ||
    (browserCard && !tryFrame && Boolean(browserPageName(savedUrl))));
  sizeActivity();
  if (localHtml) {
    editor.hidden = true;
    change.hidden = true;
    openPage.hidden = true;
    if (!frame.isConnected) viewport.append(frame);
    if (!frame.hasAttribute('srcdoc')) frame.srcdoc = localHtml;
    frame.hidden = false;
    return;
  }
  const showForm = editable && (editing || !savedUrl);
  editor.hidden = !showForm;
  const pageName = browserPageName(savedUrl);
  const showBrowserCard = browserCard && !showForm && !tryFrame && Boolean(pageName);
  const direct = offersDirect(savedUrl);
  const canLaunch = direct && presenting && !showForm && !saving;
  directLink('direct-card', canLaunch && showBrowserCard);
  directLink('direct-nav', canLaunch && !showBrowserCard);
  if (showForm) updateHint();
  if (browserCard) browserCard.hidden = !showBrowserCard;
  if (showBrowserCard) {
    element('browser-card-title').textContent = pageName;
    const link = element('browser-card-link');
    link.href = savedUrl;
    const queensHome = pageName === 'Queen’s University Belfast';
    link.hidden = canLaunch && !queensHome;
    link.textContent = queensHome ? 'Open Queen’s homepage in browser' : 'Open ' + pageName;
    const destination = element('destination');
    if (destination) destination.textContent = new URL(savedUrl).host;
    const instructions = element('card-instructions');
    if (instructions) {
      if (queensHome) {
        const launchHint = canLaunch ? 'Choose Open on slide to try the homepage directly.' : direct ? 'Start Slide Show and choose Open on slide to try the homepage directly.' : 'Choose Open in browser to visit the homepage.';
        instructions.textContent = launchHint + ' Queen’s may refuse the request. If you see a 403 error, try the link in your usual browser.';
      } else if (canLaunch) instructions.textContent = 'Open this page in the PowerPoint add-in area. The website will replace the viewer controls for this visit.';
      else if (direct) instructions.textContent = 'Start Slide Show and choose Open on slide to launch this page directly in desktop PowerPoint.';
      else instructions.textContent = 'This homepage restricts embedding. Open the page in your browser, then return to your slide.';
    }
    element('try-frame').hidden = direct;
  }
  const frameHelp = element('frame-help');
  if (frameHelp) frameHelp.hidden = !savedUrl || showForm || showBrowserCard;
  if (showForm || !savedUrl || showBrowserCard) {
    frame.remove();
  } else {
    if (frame.getAttribute('src') !== savedUrl) frame.src = savedUrl;
    frame.hidden = false;
    if (!frame.isConnected) viewport.append(frame);
  }
  change.hidden = !editable || showForm;
  cancel.hidden = !savedUrl;
  cancel.disabled = saving;
  input.disabled = saving;
  save.disabled = saving;
  openPage.hidden = !savedUrl;
  if (savedUrl) openPage.href = savedUrl;
}

function editAddress() {
  editing = true;
  error.textContent = '';
  input.value = savedUrl;
  render();
  input.focus();
  input.select();
}

element('try-frame')?.addEventListener('click', () => {
  tryFrame = true;
  render();
  status.textContent = 'Trying the page in the slide. If it stays blank, choose Open in browser.';
  openPage.focus();
});

change.addEventListener('click', editAddress);
element('request-site')?.addEventListener('focus', event => event.target.select());
element('samples')?.addEventListener('click', event => {
  const choice = event.target.closest('button[data-url]');
  if (!choice || saving) return;
  input.value = choice.dataset.url;
  error.textContent = '';
  updateHint();
  element('samples').open = false;
  input.focus();
});
cancel.addEventListener('click', () => {
  editing = false;
  error.textContent = '';
  render();
  change.focus();
});
input.addEventListener('input', () => {
  error.textContent = '';
  updateHint();
});

input.addEventListener('keydown', event => {
  if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) {
    event.preventDefault();
    save.click();
  }
});

form.addEventListener('submit', event => {
  event.preventDefault();
  if (!editable || saving) return;
  const url = addressFromEntry(input.value);
  if (!url) {
    error.textContent = input.value.trimStart().startsWith('<')
      ? 'Paste embed code with one iframe and a full HTTPS source address, or open the item and copy its HTTPS link.'
      : 'Enter the HTTPS address of the page to display. Use a link without an embedded username or password.';
    input.focus();
    return;
  }

  function complete() {
    tryFrame = false;
    savedUrl = url;
    editing = false;
    saving = false;
    error.textContent = '';
    status.textContent = preview ? 'Browser preview · changes last for this visit.' :
      (editable ? 'URL saved. Save your presentation to keep it.' : '');
    render();
    if (editable) change.focus();
  }

  if (preview) { complete(); return; }

  // Settings.set changes memory immediately. Restore the previous value if Office rejects the save.
  let previous;
  function failed() {
    saving = false;
    status.textContent = '';
    error.textContent = 'PowerPoint could not save this URL. Check that the presentation is editable and try again.';
    render();
    if (editable) input.focus();
  }
  try {
    previous = settings.get(settingKey);
    settings.set(settingKey, url);
  } catch {
    failed();
    return;
  }
  function rollBack() {
    try {
      if (previous == null) settings.remove(settingKey);
      else settings.set(settingKey, previous);
    } catch {
      failed();
      error.textContent = 'Reopen the presentation before trying again. PowerPoint could not restore the previous setting.';
      editable = false;
      render();
      showHelp(error.textContent);
      return;
    }
    failed();
  }
  saving = true;
  error.textContent = '';
  status.textContent = 'Saving the URL to this slide…';
  render();
  try {
    settings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) complete();
      else rollBack();
    });
  } catch { rollBack(); }
});

function showHelp(message) {
  status.textContent = message;
  element('help').hidden = false;
}

function connect(host, platform) {
  if (host !== Office.HostType.PowerPoint) {
    showHelp('Insert Slide Web Viewer in a PowerPoint slide to save a URL with the presentation.');
    return;
  }
  const doc = Office.context.document;
  const desktop = Office.PlatformType && (platform === Office.PlatformType.PC || platform === Office.PlatformType.Mac);
  desktopDirect = Boolean(platform && desktop && launch.get('direct') === '1');
  settings = doc.settings;
  savedUrl = pageAddress(settings.get(settingKey));
  const encoded = settings.get('qubSlideStudio.html.v1');
  if (!savedUrl && typeof encoded === 'string' && encoded.length <= 1400000) {
    try {
      const bytes = Uint8Array.from(atob(encoded), char => char.charCodeAt(0));
      localHtml = new TextDecoder('utf-8', {fatal:true}).decode(bytes);
    } catch { showHelp('The HTML file on this slide could not be read. Insert it again from QUB Slide Studio.'); return; }
  }
  input.value = savedUrl;
  function viewChanged(view) {
    element('help').hidden = true;
    presenting = view === 'read';
    editable = view === 'edit' && doc.mode !== Office.DocumentMode.ReadOnly;
    if (!editable) editing = false;
    error.textContent = '';
    status.textContent = !editable && !savedUrl && !localHtml ? 'Open the presentation in editing view to choose a webpage for this slide.' : '';
    render();
  }
  let gotEvent = false;
  doc.addHandlerAsync(Office.EventType.ActiveViewChanged, event => {
    gotEvent = true;
    viewChanged(event.activeView);
  });
  doc.getActiveViewAsync(result => {
    if (gotEvent) return;
    if (result.status === Office.AsyncResultStatus.Succeeded) viewChanged(result.value);
    else {
      render();
      showHelp('PowerPoint could not report its editing view. Reopen the presentation to change this URL.');
    }
  });
}

if (preview) {
  editable = true;
  status.textContent = 'Browser preview · changes last for this visit.';
  save.textContent = 'Preview webpage';
  render();
} else {
  const unavailable = () => showHelp('PowerPoint could not start the viewer. Check your connection and reopen the presentation.');
  const timer = setTimeout(unavailable, 15000);
  if (window.Office) {
    Office.onReady().then(info => {
      clearTimeout(timer);
      connect(info.host, info.platform);
    }).catch(() => { clearTimeout(timer); unavailable(); });
  } else {
    clearTimeout(timer);
    unavailable();
  }
}
})().catch(() => {
  document.getElementById('status').textContent = 'The viewer could not start. Reopen the presentation and check your connection.';
});
