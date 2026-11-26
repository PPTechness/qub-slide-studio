/* Shared by the on-slide viewer and the authoring task pane. */
(() => {
'use strict';
function pageAddress(value) {
  if (typeof value !== 'string' || !/^https:\/\//i.test(value.trim())) return '';
  try {
    const url = new URL(value.trim());
    if (url.protocol !== 'https:' || url.username || url.password) return '';
    if (url.origin === location.origin &&
        url.pathname.replace(/\.html$/, '') === location.pathname.replace(/\.html$/, '')) return '';
    return url.href;
  } catch {
    return '';
  }
}

function addressFromEntry(value) {
  const text = value.trim();
  if (!text.startsWith('<')) return pageAddress(text);
  // Template contents stay inert. Extract only the URL and discard the fragment.
  const template = document.createElement('template');
  template.innerHTML = text;
  const frames = template.content.querySelectorAll('iframe');
  if (frames.length !== 1 || frames[0].hasAttribute('srcdoc')) return '';
  const src = frames[0].getAttribute('src')?.trim() || '';
  return pageAddress(src.startsWith('//') ? 'https:' + src : src);
}

globalThis.SlideAddresses = Object.freeze({pageAddress, addressFromEntry});
})();
