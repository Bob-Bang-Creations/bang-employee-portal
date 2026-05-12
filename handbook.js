// ── Handbook ──────────────────────────────────────────────
// Depends on: config.js, graph.js, main.js

// Cached content of handbook.css — fetched once during initHandbook so the
// viewer can inject it into every document blob without a network round-trip.
var gHbCss = '';

// ── Group check ──────────────────────────────────────────
function checkHandbookEditor() {
  return gPost('/me/checkMemberObjects', { ids: [CFG.hbGroupId] }).then(function(d) {
    gIsHandbookEditor = d.value && d.value.indexOf(CFG.hbGroupId) >= 0;
  }).catch(function() {
    gIsHandbookEditor = false;
  });
}

// ── Fetch ────────────────────────────────────────────────
function fetchHbDrive() {
  return gGet('/sites/' + gSiteId + '/drives').then(function(d) {
    var drives = d.value || [];
    // Default SharePoint Documents library is named 'Documents' in Graph
    var drive = drives.filter(function(dr) { return dr.name === 'Documents'; })[0]
              || drives[0];
    if (!drive) throw new Error('Handbook: no document library found on this site.');
    gHbDriveId = drive.id;
  });
}

// Map a raw Graph driveItem into a handbook doc object.
function hbMapItem(item) {
  var fields = (item.listItem && item.listItem.fields) || {};
  // FileTypeTag is a multi-choice column — Graph may return an array or a
  // ';#'-delimited string depending on the SharePoint column configuration.
  var raw  = fields.FileTypeTag;
  var tags = [];
  if (Array.isArray(raw)) {
    tags = raw;
  } else if (typeof raw === 'string' && raw) {
    tags = raw.indexOf(';#') >= 0 ? raw.split(';#').filter(Boolean) : [raw];
  }
  return {
    id:      item.id,
    name:    item.name.replace(/\.html?$/i, ''),
    tags:    tags,
    version: fields._UIVersionString || '',
    updated: (item.lastModifiedDateTime || '').split('T')[0],
    webUrl:  item.webUrl
  };
}

function fetchHbDocs() {
  var pathStr = encodeURIComponent(CFG.hbFolder);
  return gGet(
    '/drives/' + gHbDriveId +
    '/root:/' + pathStr + ':/children' +
    '?$expand=listItem($expand=fields)&$top=500'
  ).then(function(d) {
    return (d.value || [])
      .filter(function(i) { return !i.folder && /\.html?$/i.test(i.name); })
      .map(hbMapItem);
  }).then(function(docs) {
    gHbDocs = docs;

    // Build sorted unique tag list from all documents
    var seen = {};
    gHbDocs.forEach(function(doc) {
      doc.tags.forEach(function(t) { if (t) seen[t] = true; });
    });
    gHbTags = Object.keys(seen).sort();
  });
}

// Fetch and cache handbook.css so the viewer can inject it into every blob
// without an extra round-trip per document open.
function fetchHbCss() {
  return fetch('./handbook.css')
    .then(function(r) { return r.ok ? r.text() : ''; })
    .then(function(css) { gHbCss = css; })
    .catch(function() { gHbCss = ''; });
}

// Wrap an HTML fragment (or full document) in a complete document shell with
// handbook.css injected inline. Blob URLs have no base path so external
// stylesheets referenced by relative URLs would not resolve — injecting the
// CSS as a <style> block avoids that entirely.
function hbWrapDoc(html) {
  var isFullDoc = /^\s*<!doctype/i.test(html) || /^\s*<html/i.test(html);
  var styleBlock = gHbCss ? '<style>\n' + gHbCss + '\n</style>' : '';

  if (isFullDoc) {
    // Full document: splice the style block before </head>, or after <head>
    if (/<\/head>/i.test(html)) return html.replace(/<\/head>/i, styleBlock + '</head>');
    if (/<head>/i.test(html))   return html.replace(/<head>/i, '<head>' + styleBlock);
    return html;
  }

  // Fragment (standard template format): build a minimal document shell
  return '<!doctype html>'
    + '<html lang="en"><head>'
    + '<meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + styleBlock
    + '</head><body>'
    + html
    + '</body></html>';
}

function initHandbook() {
  return fetchHbCss()
    .then(fetchHbDrive)
    .then(fetchHbDocs)
    .then(renderHandbook);
}

// ── Filtering ────────────────────────────────────────────
function hbFiltered() {
  var q = gHbSearch.toLowerCase().trim();
  return gHbDocs.filter(function(doc) {
    var matchQ = !q
      || doc.name.toLowerCase().indexOf(q) >= 0
      || doc.tags.some(function(t) { return t.toLowerCase().indexOf(q) >= 0; });
    var matchT = !gHbActiveTags.length
      || gHbActiveTags.some(function(at) { return doc.tags.indexOf(at) >= 0; });
    return matchQ && matchT;
  });
}

// ── Render ────────────────────────────────────────────────
function renderHandbook() {
  var meta = document.getElementById('hb-meta');
  if (meta) meta.textContent = gHbDocs.length + ' document' + (gHbDocs.length !== 1 ? 's' : '');
  renderHbTags();
  renderHbResults();
}

function renderHbTags() {
  var el = document.getElementById('hb-tags');
  if (!el) return;
  if (!gHbTags.length) {
    el.innerHTML = '<span style="font-size:13px;color:var(--g400)">No tags found.</span>';
    return;
  }
  el.innerHTML = gHbTags.map(function(t) {
    var act = gHbActiveTags.indexOf(t) >= 0;
    return '<button class="hb-tag' + (act ? ' act' : '') + '" data-tag="' + escHtml(t) + '"'
      + ' onclick="hbToggleTag(this.dataset.tag)">' + escHtml(t) + '</button>';
  }).join('');
}

function renderHbResults() {
  var el      = document.getElementById('hb-results');
  var countEl = document.getElementById('hb-count');
  if (!el) return;

  var docs = hbFiltered();
  if (countEl) countEl.textContent = docs.length + ' of ' + gHbDocs.length;

  if (!docs.length) {
    el.innerHTML = '<div class="empty">'
      + (gHbDocs.length ? 'No documents match your search.' : 'No documents found in the Handbook folder.')
      + '</div>';
    return;
  }

  el.innerHTML = '<div class="tblcard" style="margin-bottom:0">'
    + docs.map(function(doc) {
        var tagHtml = doc.tags.length
          ? doc.tags.map(function(t) {
              return '<span class="badge bbl" style="font-size:10.5px">' + escHtml(t) + '</span>';
            }).join('')
          : '<span style="font-size:11px;color:var(--g400)">Untagged</span>';

        var editBtn = gIsHandbookEditor
          ? '<button class="hb-act-btn" title="Edit in SharePoint"'
            + ' data-id="' + escHtml(doc.id) + '"'
            + ' onclick="event.stopPropagation();hbOpenEdit(this.dataset.id)">Edit</button>'
          : '';

        return '<div class="hb-row" data-id="' + escHtml(doc.id) + '"'
          + ' onclick="hbOpenViewer(this.dataset.id)">'
          + '<div class="hb-row-main">'
            + '<div class="hb-row-name">' + escHtml(doc.name) + '</div>'
            + '<div class="hb-row-tags">' + tagHtml + '</div>'
          + '</div>'
          + '<div class="hb-row-date">Updated ' + fmt(doc.updated) + '</div>'
          + '<div class="hb-row-actions">'
            + editBtn
            + '<button class="hb-act-btn" title="Copy link"'
              + ' data-id="' + escHtml(doc.id) + '"'
              + ' onclick="event.stopPropagation();hbShare(this.dataset.id)">Share</button>'
          + '</div>'
          + '</div>';
      }).join('')
    + '</div>';
}

// ── Interactions ──────────────────────────────────────────
function hbSearch(val) {
  gHbSearch = val;
  renderHbResults();
}

function hbToggleTag(tag) {
  var idx = gHbActiveTags.indexOf(tag);
  if (idx >= 0) gHbActiveTags.splice(idx, 1);
  else gHbActiveTags.push(tag);
  renderHbTags();
  renderHbResults();
}

// ── Viewer ────────────────────────────────────────────────
function hbOpenViewer(docId) {
  var doc = null;
  for (var i = 0; i < gHbDocs.length; i++) {
    if (gHbDocs[i].id === docId) { doc = gHbDocs[i]; break; }
  }
  if (!doc) return;

  gHbViewing = docId;
  document.getElementById('hb-viewer-title').textContent = doc.name;

  var metaEl = document.getElementById('hb-viewer-meta');
  if (metaEl) {
    var metaHtml = doc.tags.map(function(t) {
      return '<span class="badge bbl" style="font-size:10.5px">' + escHtml(t) + '</span>';
    }).join('');
    if (doc.version) {
      metaHtml += '<span class="badge" style="font-size:10.5px;background:var(--g50);color:var(--g600);padding:2px 8px;border-radius:20px">Version ' + escHtml(doc.version) + '</span>';
    }
    metaEl.innerHTML = metaHtml;
  }

  // Reset frame and show loading spinner
  var frame   = document.getElementById('hb-viewer-frame');
  var loading = document.getElementById('hb-viewer-loading');
  frame.src              = 'about:blank';
  frame.style.display    = 'none';
  loading.style.display  = 'flex';

  // Edit button — Handbook Editors only
  var editBtn = document.getElementById('hb-viewer-editbtn');
  if (editBtn) {
    editBtn.style.display = gIsHandbookEditor ? '' : 'none';
    editBtn.dataset.id    = docId;
  }

  // Show overlay and lock body scroll
  document.getElementById('hb-viewer').style.display = 'flex';
  document.body.style.overflow = 'hidden';

  // Write deep-link hash
  history.replaceState(null, '', location.pathname + location.search + '#handbook/' + docId);

  // Swap frame in once content is loaded
  frame.onload = function() {
    if (frame.src === 'about:blank') return;
    loading.style.display = 'none';
    frame.style.display   = 'block';
  };

  // Fetch fresh pre-auth download URL then load HTML as a blob
  gGet('/drives/' + gHbDriveId + '/items/' + docId)
    .then(function(item) {
      var dlUrl = item['@microsoft.graph.downloadUrl'];
      if (!dlUrl) throw new Error('No download URL available for this document.');
      return fetch(dlUrl);
    })
    .then(function(r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' fetching document.');
      return r.text();
    })
    .then(function(html) {
      if (gHbViewing !== docId) return; // viewer was closed before load finished
      if (doc.version) html = html.replace('{{version}}', 'Version ' + doc.version);
      if (gHbBlobUrl) URL.revokeObjectURL(gHbBlobUrl);
      gHbBlobUrl = URL.createObjectURL(new Blob([hbWrapDoc(html)], { type: 'text/html' }));
      frame.src  = gHbBlobUrl;
    })
    .catch(function(e) {
      if (gHbViewing !== docId) return;
      hbShowToast('Could not load document: ' + (e.message || String(e)), true);
      closeHandbookViewer();
    });
}

function closeHandbookViewer() {
  gHbViewing = null;
  var frame   = document.getElementById('hb-viewer-frame');
  var loading = document.getElementById('hb-viewer-loading');
  frame.onload           = null;
  frame.src              = 'about:blank';
  frame.style.display    = 'none';
  loading.style.display  = 'flex';
  document.getElementById('hb-viewer').style.display = 'none';
  document.body.style.overflow = '';
  if (gHbBlobUrl) { URL.revokeObjectURL(gHbBlobUrl); gHbBlobUrl = null; }
  history.replaceState(null, '', location.pathname + location.search);
}

function hbOpenEdit(docId) {
  var id = docId || gHbViewing;
  for (var i = 0; i < gHbDocs.length; i++) {
    if (gHbDocs[i].id === id) {
      window.open(gHbDocs[i].webUrl, '_blank', 'noopener,noreferrer');
      return;
    }
  }
}

function hbShare(docId) {
  var id  = docId || gHbViewing;
  var url = location.origin + location.pathname + '#handbook/' + id;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(url)
      .then(function()  { hbShowToast('Link copied to clipboard!'); })
      .catch(function() { hbShowToast('Could not copy link.', true); });
  } else {
    // Fallback for older browsers
    var ta = document.createElement('textarea');
    ta.value = url;
    ta.style.cssText = 'position:fixed;top:-9999px;left:-9999px';
    document.body.appendChild(ta);
    ta.select();
    try   { document.execCommand('copy'); hbShowToast('Link copied to clipboard!'); }
    catch (e) { hbShowToast('Could not copy link.', true); }
    document.body.removeChild(ta);
  }
}

// ── Deep-link handler ─────────────────────────────────────
// Called once after bootstrap completes and the portal is visible.
function hbHandleDeepLink() {
  var m = location.hash.match(/^#handbook\/(.+)$/);
  if (!m) return;
  var docId = m[1];

  // Switch to the Handbook tab
  var tabs = document.querySelectorAll('.tab');
  for (var i = 0; i < tabs.length; i++) {
    if ((tabs[i].getAttribute('onclick') || '').indexOf("'handbook'") >= 0) {
      switchTab('handbook', tabs[i]);
      break;
    }
  }
  // Docs are already loaded at this point (initHandbook ran during bootstrap)
  hbOpenViewer(docId);
}

// ── Esc key closes viewer ─────────────────────────────────
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape' && gHbViewing) closeHandbookViewer();
});

// ── Toast ─────────────────────────────────────────────────
function hbShowToast(msg, isErr) {
  var el = document.getElementById('hb-toast');
  var tx = document.getElementById('hb-toasttxt');
  if (!el || !tx) return;
  tx.textContent = msg;
  el.classList.toggle('err', !!isErr);
  el.classList.add('show');
  setTimeout(function() { el.classList.remove('show'); }, 4000);
}
