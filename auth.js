// ── Auth ─────────────────────────────────────────────────
// Depends on: config.js (CFG, SCOPES, gMsal, gUser), graph.js

function doSignIn() {
  document.getElementById('sin-err').style.display = 'none';
  gMsal.loginRedirect({ scopes: SCOPES }).catch(function(e) {
    var el = document.getElementById('sin-err');
    el.textContent = 'Sign-in failed: ' + (e.message || String(e));
    el.style.display = 'block';
  });
}

function doSignOut() {
  gMsal.logoutRedirect();
}

function getToken() {
  var acct = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  if (!acct) return Promise.reject(new Error('Not signed in'));
  return gMsal.acquireTokenSilent({ scopes: SCOPES, account: acct }).then(function(r) { return r.accessToken; });
}

// ── Group membership checks ──────────────────────────────
function checkHrAdmin() {
  return gPost('/me/checkMemberObjects', { ids: [CFG.hrGroupId] }).then(function(d) {
    gIsHrAdmin = d.value && d.value.indexOf(CFG.hrGroupId) >= 0;
  }).catch(function() {
    gIsHrAdmin = false;
  });
}

function checkAccMgr() {
  return gPost('/me/checkMemberObjects', { ids: [CFG.accGroupId] }).then(function(d) {
    gIsAccMgr = d.value && d.value.indexOf(CFG.accGroupId) >= 0;
  }).catch(function() {
    gIsAccMgr = false;
  });
}

// ── Bootstrap ────────────────────────────────────────────
function bootstrap() {
  setLoading('Connecting to Microsoft 365…');
  var acct  = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  if (!acct) { showScreen('s-signin'); return; }

  var email = (acct.username || '').toLowerCase();
  var name  = (acct.name || acct.username).replace(/\s*\|\s*Bang Creations\s*/i, '');
  gUser = {
    name:  name,
    email: email,
    ini:   name.split(' ').map(function(p) { return p[0]; }).join('').slice(0, 2).toUpperCase()
  };

  setLoading('Connecting to SharePoint…');
  getSiteId()
  .then(function(id) {
    gSiteId = id;
    setLoading('Loading lists…');
    return getListId(id, CFG.listTasks);
  })
  .then(function(id) {
    gTLId = id;
    return getListId(gSiteId, CFG.listExpenses);
  })
  .then(function(id) {
    gELId = id;
    setLoading('Checking permissions…');
    return checkHrAdmin();
  })
  .then(function() {
    return checkAccMgr();
  })
  .then(function() {
    return checkHandbookEditor();
  })
  .then(function() {
    setLoading('Loading users…');
    return fetchUsers();
  })
  .then(function() {
    setLoading('Loading tasks…');
    return fetchTasks();
  })
  .then(function() {
    return fetchExpenses();
  })
  .then(function() {
    setLoading('Loading handbook…');
    return initHandbook();
  })
  .then(function() {
    document.getElementById('uname').textContent = gUser.name;
    document.getElementById('uav').textContent   = gUser.ini;
    document.getElementById('sobtn').onclick      = doSignOut;
    showScreen('portal');
    applyHrAccess();
    applyAccAccess();
    renderMy();
    renderManage();
    renderMyExpenses();
    renderAccMgrTable();
    hbHandleDeepLink();
  })
  .catch(function(e) {
    showFatalError(e.message || String(e));
    console.error(e);
  });
}

// ── Startup: load MSAL then initialise ───────────────────
setLoading('Loading…');

var msalScript   = document.createElement('script');
msalScript.src   = './msal-browser.min.js';

msalScript.onload = function() {
  var msalCfg = {
    auth: {
      clientId:    CFG.clientId,
      authority:   'https://login.microsoftonline.com/' + CFG.tenantId,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: 'memory', storeAuthStateInCookie: true }
  };
  gMsal = new msal.PublicClientApplication(msalCfg);

  gMsal.initialize().then(function() {
    document.getElementById('sinbtn').onclick = doSignIn;
    return gMsal.handleRedirectPromise();
  }).then(function(res) {
    if (res && res.account) {
      gMsal.setActiveAccount(res.account);
      bootstrap();
      return;
    }
    var accts = gMsal.getAllAccounts();
    if (accts.length > 0) {
      gMsal.setActiveAccount(accts[0]);
      bootstrap();
      return;
    }
    // Memory cache loses accounts on page refresh — attempt silent SSO before showing sign-in
    return gMsal.ssoSilent({ scopes: SCOPES }).then(function(r) {
      gMsal.setActiveAccount(r.account);
      bootstrap();
    }).catch(function() {
      showScreen('s-signin');
    });
  }).catch(function(e) {
    showFatalError('Initialisation error: ' + (e.message || String(e)));
  });
};

msalScript.onerror = function() {
  showFatalError('Could not load msal-browser.min.js — please make sure it is uploaded to your GitHub repository alongside this file.');
};

document.body.appendChild(msalScript);
