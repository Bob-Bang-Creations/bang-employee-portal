// ── App-level helpers ─────────────────────────────────────
// Depends on: config.js, hr-tasks.js, expenses.js

function showScreen(id) {
  var ids = ['s-loading', 's-signin', 's-error', 'portal'];
  for (var i = 0; i < ids.length; i++) {
    var el = document.getElementById(ids[i]);
    if (!el) continue;
    el.classList.remove('active');
    if (ids[i] === 'portal') el.style.display = 'none';
  }
  var t = document.getElementById(id);
  if (!t) return;
  if (id === 'portal') { t.style.display = 'flex'; t.classList.add('active'); }
  else t.classList.add('active');
}

function setLoading(msg) {
  document.getElementById('ltxt').textContent = msg;
  showScreen('s-loading');
}

function showFatalError(msg) {
  document.getElementById('errmsg').textContent = msg;
  showScreen('s-error');
}

function switchTab(name, btn) {
  // If leaving the Handbook tab while viewer is open, close it cleanly
  if (name !== 'handbook' && gHbViewing) closeHandbookViewer();

  var panels = document.querySelectorAll('.panel');
  for (var i = 0; i < panels.length; i++) panels[i].classList.remove('active');
  var tabs = document.querySelectorAll('.tab');
  for (var i = 0; i < tabs.length; i++) tabs[i].classList.remove('active');
  document.getElementById('panel-' + name).classList.add('active');
  if (btn) btn.classList.add('active');
  if (name === 'expenses') renderMyExpenses();
}

function applyHrAccess() {
  var section = document.getElementById('hr-admin-section');
  if (section) section.style.display = gIsHrAdmin ? '' : 'none';
}

function applyAccAccess() {
  var section = document.getElementById('acc-admin-section');
  if (section) section.style.display = gIsAccMgr ? '' : 'none';
}
