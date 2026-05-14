// ── Graph helpers ────────────────────────────────────────
// Depends on: config.js (CFG, SCOPES, gMsal, gUsers, gSiteId)

function gGet(url) {
  return getToken().then(function(tok) {
    return fetch('https://graph.microsoft.com/v1.0' + url, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json' }
    }).then(function(r) {
      if (!r.ok) return r.text().then(function(t) { throw new Error('GET ' + url + ' => ' + r.status + ' ' + t); });
      return r.json();
    });
  });
}

function gPost(url, body) {
  return getToken().then(function(tok) {
    return fetch('https://graph.microsoft.com/v1.0' + url, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function(r) {
      if (!r.ok) return r.text().then(function(t) { throw new Error('POST ' + url + ' => ' + r.status + ' ' + t); });
      return r.json();
    });
  });
}

function gPatch(url, body) {
  return getToken().then(function(tok) {
    return fetch('https://graph.microsoft.com/v1.0' + url, {
      method: 'PATCH',
      headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function(r) {
      if (!r.ok) return r.text().then(function(t) { throw new Error('PATCH ' + url + ' => ' + r.status + ' ' + t); });
    });
  });
}

function gDel(url) {
  return getToken().then(function(tok) {
    return fetch('https://graph.microsoft.com/v1.0' + url, {
      method: 'DELETE',
      headers: { Authorization: 'Bearer ' + tok }
    }).then(function(r) {
      if (!r.ok) throw new Error('DELETE ' + url + ' => ' + r.status);
    });
  });
}

// ── SharePoint site + list resolution ───────────────────
function getSiteId() {
  var host = CFG.tenant + '.sharepoint.com';
  var path = CFG.siteUrl.replace('https://' + host, '');
  return gGet('/sites/' + host + ':' + path).then(function(d) { return d.id; });
}

function getListId(sid, name) {
  return gGet('/sites/' + sid + '/lists?$filter=displayName eq \'' + name + '\'').then(function(d) {
    if (d.value && d.value.length > 0) return d.value[0].id;
    throw new Error('SharePoint list "' + name + '" not found. Please create it manually — see setup instructions.');
  });
}

// ── Fetch org users from Graph (paginated) ───────────────
function fetchUsers() {
  return fetchUsersPage('/users?$select=displayName,mail,userPrincipalName&$top=999', []);
}

function fetchUsersPage(url, acc) {
  var fullUrl = url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url;
  return getToken().then(function(tok) {
    return fetch(fullUrl, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json' }
    }).then(function(r) {
      if (!r.ok) return r.text().then(function(t) { throw new Error('GET users => ' + r.status + ' ' + t); });
      return r.json();
    });
  }).then(function(d) {
    var page = (d.value || [])
      .filter(function(u) {
        var upn = u.userPrincipalName || '';
        return u.displayName
          && (u.mail || u.userPrincipalName)
          && upn.indexOf('#EXT#') === -1;
      })
      .map(function(u) { return { name: u.displayName, email: (u.mail || u.userPrincipalName).toLowerCase() }; });
    var all = acc.concat(page);
    if (d['@odata.nextLink']) return fetchUsersPage(d['@odata.nextLink'], all);
    gUsers = all.sort(function(a, b) { return a.name.localeCompare(b.name); });
    populateUserDropdown();
  });
}

// ── Utilities ────────────────────────────────────────────
function today()     { return new Date().toISOString().split('T')[0]; }
function fmt(d)      { if (!d) return '—'; var p = d.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
function ini(n)      { return n.split(' ').map(function(p) { return p[0]; }).join(''); }
function cleanName(n){ return (n || '').split(' | ')[0].trim(); }
function escHtml(s)  { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function safeUrl(u)  { return /^https?:\/\//i.test(u) ? u : (u ? 'https://' + u : '#'); }
function tryParse(s, def) { try { return JSON.parse(s || ''); } catch(e) { return def; } }

function nameByEmail(email) {
  var e = email.toLowerCase();
  for (var i = 0; i < gUsers.length; i++) {
    if (gUsers[i].email === e) return gUsers[i].name;
  }
  return email;
}
