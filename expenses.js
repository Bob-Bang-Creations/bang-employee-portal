// ── Expenses ─────────────────────────────────────────────
// Depends on: config.js, graph.js

// ── Fetch ────────────────────────────────────────────────
function fetchExpenses() {
  return gGet('/sites/' + gSiteId + '/lists/' + gELId + '/items?$expand=fields&$top=500').then(function(d) {
    gExpenses = d.value.map(function(i) {
      var f = i.fields;
      return {
        id:          i.id,
        title:       f.Title || '',
        date:        (f.expense_date || f.date_of_expense || '').split('T')[0],
        client:      f.client_number || '',
        project:     f.project_number || '',
        amount:      parseFloat(f.amount) || 0,
        amountExVat: parseFloat(f.amount_exVAT) || 0,
        amountVat:   parseFloat(f.amount_VAT) || 0,
        incVat:      !!f.inc_VAT,
        vatRate:     parseFloat(f.VAT_rate) || 0,
        processed:      !!f.is_processed,
        hasAttachment:  !!f.Attachments,
        person:         (f.person || '').toLowerCase()
      };
    });
  });
}

function fetchMileage() {
  return gGet('/sites/' + gSiteId + '/lists/' + gMLId + '/items?$expand=fields&$top=500').then(function(d) {
    gMileage = d.value.map(function(i) {
      var f = i.fields;
      return {
        id:       i.id,
        title:    f.Title || '',
        date:     (f.travel_date || '').split('T')[0],
        client:   f.client_number || '',
        project:  f.project_number || '',
        distance:  parseFloat(f.distance) || 0,
        rate:      parseFloat(f.rate) || 0.45,
        processed: !!f.is_processed,
        person:    (f.person || '').toLowerCase()
      };
    });
  });
}

// ── UI helpers ───────────────────────────────────────────
function setExpType(type) {
  gExpType = type;
  document.getElementById('exp-std-form').style.display = type === 'standard' ? '' : 'none';
  document.getElementById('exp-mil-form').style.display = type === 'mileage'  ? '' : 'none';
  var stdBtn = document.getElementById('exp-type-std');
  var milBtn = document.getElementById('exp-type-mil');
  stdBtn.style.background = type === 'standard' ? 'var(--or)' : 'transparent';
  stdBtn.style.color      = type === 'standard' ? 'var(--w)'  : 'var(--or)';
  milBtn.style.background = type === 'mileage'  ? 'var(--or)' : 'transparent';
  milBtn.style.color      = type === 'mileage'  ? 'var(--w)'  : 'var(--or)';
}

function toggleVatRate() {
  var has = document.getElementById('es-hasvat').checked;
  if (!has) {
    document.getElementById('es-vatrate').value = '0';
  } else {
    var cur = parseFloat(document.getElementById('es-vatrate').value) || 0;
    if (cur === 0) document.getElementById('es-vatrate').value = '20';
  }
  recalcVat();
}

function recalcVat() {
  var amtStr = document.getElementById('es-amount').value;
  var amt    = parseFloat(amtStr) || 0;
  var rateEl = document.getElementById('es-vatrate');
  var cbEl   = document.getElementById('es-hasvat');
  var rate   = parseFloat(rateEl.value) || 0;
  var prev   = document.getElementById('es-vat-preview');

  if (rate === 0 && cbEl.checked) cbEl.checked = false;
  if (rate > 0 && !cbEl.checked)  cbEl.checked = true;

  var has = cbEl.checked;
  if (!amt || !has || rate === 0) { prev.textContent = ''; return; }
  var exVat  = amt / (1 + rate / 100);
  var vatAmt = amt - exVat;
  prev.textContent = 'Ex-VAT: £' + exVat.toFixed(2) + '  |  VAT (' + rate + '%): £' + vatAmt.toFixed(2);
}

function recalcMileage() {
  var dist = parseFloat(document.getElementById('em-distance').value) || 0;
  var rate = parseFloat(document.getElementById('em-rate').value) || 0.45;
  var prev = document.getElementById('em-preview');
  if (!dist) { prev.textContent = ''; return; }
  prev.textContent = 'Total: £' + (dist * rate).toFixed(2);
}

function expShowToast(msg, isErr) {
  var el = document.getElementById('exp-toast'), tx = document.getElementById('exp-toasttxt');
  tx.textContent = msg; el.classList.toggle('err', !!isErr); el.classList.add('show');
  setTimeout(function() { el.classList.remove('show'); }, 4000);
}

// ── Add / update standard expense ────────────────────────
function addStdExpense() {
  var client  = document.getElementById('es-client').value.trim();
  var project = document.getElementById('es-project').value.trim();
  var date    = document.getElementById('es-date').value;
  var title   = document.getElementById('es-title').value.trim();
  var amtStr  = document.getElementById('es-amount').value.trim();
  var hasVat  = document.getElementById('es-hasvat').checked;
  var vatRate = parseFloat(document.getElementById('es-vatrate').value) || 20;
  var receipt = document.getElementById('es-receipt').files[0];

  if (!client)  { alert('Please enter a client number.'); return; }
  if (!project) { alert('Please enter a project number.'); return; }
  if (!date)    { alert('Please select a date.'); return; }
  if (!title)   { alert('Please enter a description.'); return; }
  var amt = parseFloat(amtStr);
  if (isNaN(amt) || amt <= 0) { alert('Please enter a valid amount.'); return; }

  var exVat  = hasVat ? amt / (1 + vatRate / 100) : amt;
  var vatAmt = hasVat ? amt - exVat : 0;
  var btn    = document.getElementById('es-addbtn');

  // ── Edit mode ─────────────────────────────────────────
  if (gEditingExpId && gEditingExpType === 'standard') {
    btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Updating…';
    var editId = gEditingExpId;
    var p = gPatch(
      '/sites/' + gSiteId + '/lists/' + gELId + '/items/' + editId + '/fields',
      { Title: title, expense_date: date, client_number: client, project_number: project,
        amount: amt.toFixed(2), inc_VAT: hasVat, VAT_rate: hasVat ? vatRate.toFixed(2) : '0',
        amount_exVAT: exVat.toFixed(2), amount_VAT: vatAmt.toFixed(2) }
    );
    if (receipt) {
      p = p.then(function() { return uploadReceipt(editId, receipt); });
    }
    p.then(function() {
      for (var i = 0; i < gExpenses.length; i++) {
        if (gExpenses[i].id === editId) {
          var e = gExpenses[i];
          e.title = title; e.date = date; e.client = client; e.project = project;
          e.amount = amt; e.amountExVat = exVat; e.amountVat = vatAmt;
          e.incVat = hasVat; e.vatRate = vatRate;
          break;
        }
      }
      resetStdForm(); gEditingExpId = null; gEditingExpType = null;
      btn.textContent = 'Log expense';
      renderMyExpenses(); renderAccMgrTable();
      expShowToast('Expense updated!');
    }).catch(function(e) {
      expShowToast('Error: ' + e.message, true);
    }).then(function() { btn.disabled = false; });
    return;
  }

  // ── Add mode ──────────────────────────────────────────
  if (!receipt) { alert('Please attach a receipt (PDF or image).'); return; }
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

  gPost('/sites/' + gSiteId + '/lists/' + gELId + '/items', {
    fields: {
      Title: title, expense_date: date, client_number: client, project_number: project,
      amount: amt.toFixed(2), inc_VAT: hasVat, VAT_rate: hasVat ? vatRate.toFixed(2) : '0',
      amount_exVAT: exVat.toFixed(2), amount_VAT: vatAmt.toFixed(2),
      is_processed: false, person: gUser.email
    }
  }).then(function(r) {
    return uploadReceipt(r.id, receipt).then(function() { return r; });
  }).then(function(r) {
    gExpenses.push({
      id: r.id, title: title, date: date, client: client, project: project,
      amount: amt, amountExVat: exVat, amountVat: vatAmt,
      incVat: hasVat, vatRate: vatRate, processed: false, person: gUser.email
    });
    resetStdForm();
    renderMyExpenses(); renderAccMgrTable();
    expShowToast('Expense logged successfully!');
  }).catch(function(e) {
    expShowToast('Error: ' + e.message, true);
  }).then(function() {
    btn.disabled = false; btn.textContent = 'Log expense';
  });
}

function uploadReceipt(itemId, file) {
  var siteHost = CFG.tenant + '.sharepoint.com';
  var spScope  = 'https://' + siteHost + '/.default';
  var acct     = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  return gMsal.acquireTokenSilent({ scopes: [spScope], account: acct }).then(function(spRes) {
    var tok      = spRes.accessToken;
    var sitePath = CFG.siteUrl.replace('https://' + siteHost, '');
    var url      = 'https://' + siteHost + sitePath
      + '/_api/web/lists/getbytitle(\'' + CFG.listExpenses + '\')/items(' + itemId + ')/AttachmentFiles/add(FileName=\'' + encodeURIComponent(file.name) + '\')';
    return new Promise(function(resolve, reject) {
      var reader = new FileReader();
      reader.onload = function(e) {
        fetch(url, {
          method: 'POST',
          headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/octet-stream' },
          body: e.target.result
        }).then(function(res) {
          if (!res.ok) return res.text().then(function(t) { reject(new Error('Attachment upload failed: ' + t)); });
          resolve();
        }).catch(reject);
      };
      reader.onerror = function() { reject(new Error('Failed to read receipt file.')); };
      reader.readAsArrayBuffer(file);
    });
  });
}

function resetStdForm() {
  document.getElementById('es-client').value          = '';
  document.getElementById('es-project').value         = '';
  document.getElementById('es-date').value            = '';
  document.getElementById('es-title').value           = '';
  document.getElementById('es-amount').value          = '';
  document.getElementById('es-hasvat').checked        = true;
  document.getElementById('es-vatrate').value         = '20';
  document.getElementById('es-vat-preview').textContent = '';
  document.getElementById('es-receipt').value         = '';
}

// ── Add / update mileage expense ─────────────────────────
function addMilExpense() {
  var client  = document.getElementById('em-client').value.trim();
  var project = document.getElementById('em-project').value.trim();
  var date    = document.getElementById('em-date').value;
  var title   = document.getElementById('em-title').value.trim();
  var distStr = document.getElementById('em-distance').value.trim();
  var rateStr = document.getElementById('em-rate').value.trim();

  if (!client)  { alert('Please enter a client number.'); return; }
  if (!project) { alert('Please enter a project number.'); return; }
  if (!date)    { alert('Please select a date.'); return; }
  if (!title)   { alert('Please enter a purpose.'); return; }
  var dist = parseFloat(distStr);
  var rate = parseFloat(rateStr) || 0.45;
  if (isNaN(dist) || dist <= 0) { alert('Please enter a valid distance.'); return; }

  var btn = document.getElementById('em-addbtn');

  // ── Edit mode ─────────────────────────────────────────
  if (gEditingExpId && gEditingExpType === 'mileage') {
    btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Updating…';
    var editId = gEditingExpId;
    gPatch(
      '/sites/' + gSiteId + '/lists/' + gMLId + '/items/' + editId + '/fields',
      { Title: title, travel_date: date, client_number: client, project_number: project,
        distance: dist.toString(), rate: rate.toFixed(4) }
    ).then(function() {
      for (var i = 0; i < gMileage.length; i++) {
        if (gMileage[i].id === editId) {
          var m = gMileage[i];
          m.title = title; m.date = date; m.client = client;
          m.project = project; m.distance = dist; m.rate = rate;
          break;
        }
      }
      resetMilForm(); gEditingExpId = null; gEditingExpType = null;
      btn.textContent = 'Log mileage';
      renderMyExpenses(); renderAccMgrTable();
      expShowToast('Mileage updated!');
    }).catch(function(e) {
      expShowToast('Error: ' + e.message, true);
    }).then(function() { btn.disabled = false; });
    return;
  }

  // ── Add mode ──────────────────────────────────────────
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

  gPost('/sites/' + gSiteId + '/lists/' + gMLId + '/items', {
    fields: {
      Title: title, travel_date: date, client_number: client, project_number: project,
      distance: dist.toString(), rate: rate.toFixed(4), person: gUser.email
    }
  }).then(function(r) {
    gMileage.push({ id: r.id, title: title, date: date, client: client, project: project, distance: dist, rate: rate, person: gUser.email });
    resetMilForm();
    renderMyExpenses(); renderAccMgrTable();
    expShowToast('Mileage logged successfully!');
  }).catch(function(e) {
    expShowToast('Error: ' + e.message, true);
  }).then(function() {
    btn.disabled = false; btn.textContent = 'Log mileage';
  });
}

function resetMilForm() {
  document.getElementById('em-client').value        = '';
  document.getElementById('em-project').value       = '';
  document.getElementById('em-date').value          = '';
  document.getElementById('em-title').value         = '';
  document.getElementById('em-distance').value      = '';
  document.getElementById('em-rate').value          = '0.45';
  document.getElementById('em-preview').textContent = '';
}

// ── Edit / delete ─────────────────────────────────────────
function expEdit(type, id) {
  var r = null, i;
  if (type === 'standard') {
    for (i = 0; i < gExpenses.length; i++) { if (gExpenses[i].id === id) { r = gExpenses[i]; break; } }
    if (!r) return;
    setExpType('standard');
    document.getElementById('es-client').value   = r.client;
    document.getElementById('es-project').value  = r.project;
    document.getElementById('es-date').value     = r.date;
    document.getElementById('es-title').value    = r.title;
    document.getElementById('es-amount').value   = r.amount.toFixed(2);
    document.getElementById('es-hasvat').checked = r.incVat;
    document.getElementById('es-vatrate').value  = r.vatRate.toString();
    recalcVat();
    document.getElementById('es-addbtn').textContent = 'Update expense';
  } else {
    for (i = 0; i < gMileage.length; i++) { if (gMileage[i].id === id) { r = gMileage[i]; break; } }
    if (!r) return;
    setExpType('mileage');
    document.getElementById('em-client').value   = r.client;
    document.getElementById('em-project').value  = r.project;
    document.getElementById('em-date').value     = r.date;
    document.getElementById('em-title').value    = r.title;
    document.getElementById('em-distance').value = r.distance.toString();
    document.getElementById('em-rate').value     = r.rate.toString();
    recalcMileage();
    document.getElementById('em-addbtn').textContent = 'Update mileage';
  }
  gEditingExpId   = id;
  gEditingExpType = type;
  var logDivider = document.querySelector('#panel-expenses .my-divider');
  if (logDivider) logDivider.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function expDelete(type, id) {
  if (!confirm('Delete this expense? This cannot be undone.')) return;
  var listId = type === 'standard' ? gELId : gMLId;
  gDel('/sites/' + gSiteId + '/lists/' + listId + '/items/' + id)
    .then(function() {
      if (type === 'standard') {
        gExpenses = gExpenses.filter(function(e) { return e.id !== id; });
      } else {
        gMileage = gMileage.filter(function(m) { return m.id !== id; });
      }
      if (gEditingExpId === id) {
        gEditingExpId = null; gEditingExpType = null;
        if (type === 'standard') { resetStdForm(); document.getElementById('es-addbtn').textContent = 'Log expense'; }
        else                     { resetMilForm(); document.getElementById('em-addbtn').textContent = 'Log mileage'; }
      }
      renderMyExpenses(); renderAccMgrTable();
      expShowToast('Expense deleted.');
    })
    .catch(function(e) { expShowToast('Error: ' + e.message, true); });
}

// ── Render: My expenses ──────────────────────────────────
function renderMyExpenses() {
  var email = gUser ? gUser.email : '';
  var myStd = gExpenses.filter(function(e) { return e.person === email; });
  var myMil = gMileage.filter(function(m)  { return m.person === email; });

  var all = [];
  myStd.forEach(function(e) { all.push({ type: 'standard', rec: e, date: e.date }); });
  myMil.forEach(function(m) { all.push({ type: 'mileage',  rec: m, date: m.date }); });
  all.sort(function(a, b) { return b.date.localeCompare(a.date); });

  var el = document.getElementById('my-expenses-list');
  if (!all.length) { el.innerHTML = '<div class="empty">No expenses logged yet.</div>'; return; }

  el.innerHTML = all.map(function(item) {
    var r       = item.rec;
    var canEdit = item.type === 'mileage' || !r.processed;
    var actions = canEdit
      ? '<div style="display:flex;gap:5px;margin-top:5px;justify-content:flex-end">'
          + '<button class="rbtn" style="font-size:10px;padding:2px 10px"'
            + ' onclick="expEdit(\'' + item.type + '\',\'' + escHtml(r.id) + '\')">Edit</button>'
          + '<button class="rbtn-del"'
            + ' onclick="expDelete(\'' + item.type + '\',\'' + escHtml(r.id) + '\')">Delete</button>'
        + '</div>'
      : '';

    if (item.type === 'standard') {
      var total  = '£' + r.amount.toFixed(2);
      var detail = r.incVat
        ? '<span class="badge bgy" style="font-size:10px">Inc. VAT ' + r.vatRate + '%</span> Ex-VAT: £' + r.amountExVat.toFixed(2) + ' | VAT: £' + r.amountVat.toFixed(2)
        : '<span class="badge bgy" style="font-size:10px">No VAT</span>';
      var statusBadge = r.processed
        ? '<span class="badge bgr">&#10003; Processed</span>'
        : '<span class="badge bor">Pending</span>';
      return '<div class="tc">'
        + '<div class="ttop">'
          + '<div style="flex:1;min-width:0">'
            + '<div class="tn">' + escHtml(r.title) + '</div>'
            + '<div class="tmeta">'
              + '<span class="badge bbl">Standard</span>'
              + statusBadge
              + '<span style="font-size:11.5px;color:var(--g400)">' + fmt(r.date) + '</span>'
              + '<span style="font-size:11.5px;color:var(--g600)">Client: ' + escHtml(r.client) + ' | Project: ' + escHtml(r.project) + '</span>'
            + '</div>'
            + '<div class="tdesc">' + detail + '</div>'
          + '</div>'
          + '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:0;flex-shrink:0">'
            + '<div style="font-size:16px;font-weight:600;color:var(--or);white-space:nowrap">' + total + '</div>'
            + actions
          + '</div>'
        + '</div>'
        + '</div>';
    } else {
      var tot = r.distance * r.rate;
      return '<div class="tc">'
        + '<div class="ttop">'
          + '<div style="flex:1;min-width:0">'
            + '<div class="tn">' + escHtml(r.title) + '</div>'
            + '<div class="tmeta">'
              + '<span class="badge bbl">Mileage</span>'
              + '<span style="font-size:11.5px;color:var(--g400)">' + fmt(r.date) + '</span>'
              + '<span style="font-size:11.5px;color:var(--g600)">Client: ' + escHtml(r.client) + ' | Project: ' + escHtml(r.project) + '</span>'
            + '</div>'
            + '<div class="tdesc">' + r.distance + ' miles @ £' + r.rate.toFixed(4) + '/mi</div>'
          + '</div>'
          + '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:0;flex-shrink:0">'
            + '<div style="font-size:16px;font-weight:600;color:var(--or);white-space:nowrap">£' + tot.toFixed(2) + '</div>'
            + actions
          + '</div>'
        + '</div>'
        + '</div>';
    }
  }).join('');
}

// ── Render: Accounts Manager filters ────────────────────
function renderAccTypeFilters() {
  var el = document.getElementById('acc-type-filters');
  if (!el) return;
  var types = [
    { key: 'all',      label: 'All' },
    { key: 'standard', label: 'Standard' },
    { key: 'mileage',  label: 'Mileage' }
  ];
  el.innerHTML = '<span style="font-size:12px;font-weight:600;color:var(--g400);text-transform:uppercase;letter-spacing:.06em">Type:</span>'
    + types.map(function(t) {
        return '<button class="rbtn' + (gAccFilter === t.key ? ' act' : '') + '"'
          + ' onclick="setAccFilter(\'' + t.key + '\')"'
          + ' style="font-size:11px;padding:4px 14px">' + t.label + '</button>';
      }).join('');
}

function renderAccPersonFilters() {
  var el = document.getElementById('acc-person-filters');
  if (!el) return;
  var seen = {};
  gExpenses.forEach(function(e) { if (e.person) seen[e.person] = true; });
  gMileage.forEach(function(m)  { if (m.person) seen[m.person] = true; });
  var persons = Object.keys(seen).map(function(email) {
    return { email: email, name: cleanName(nameByEmail(email)) };
  }).sort(function(a, b) { return a.name.localeCompare(b.name); });

  if (!persons.length) { el.innerHTML = ''; return; }
  el.innerHTML = '<span style="font-size:12px;font-weight:600;color:var(--g400);text-transform:uppercase;letter-spacing:.06em">Person:</span>'
    + '<button class="rbtn' + (!gAccFilterPerson ? ' act' : '') + '"'
      + ' onclick="setAccFilterPerson(null)" style="font-size:11px;padding:4px 14px">All</button>'
    + persons.map(function(p) {
        return '<button class="rbtn' + (gAccFilterPerson === p.email ? ' act' : '') + '"'
          + ' onclick="setAccFilterPerson(\'' + escHtml(p.email) + '\')"'
          + ' style="font-size:11px;padding:4px 14px">' + escHtml(p.name) + '</button>';
      }).join('');
}

function renderAccStateFilters() {
  var el = document.getElementById('acc-state-filters');
  if (!el) return;
  var states = [
    { key: 'all',         label: 'All' },
    { key: 'unprocessed', label: 'Unprocessed' },
    { key: 'processed',   label: 'Processed' }
  ];
  el.innerHTML = '<span style="font-size:12px;font-weight:600;color:var(--g400);text-transform:uppercase;letter-spacing:.06em">State:</span>'
    + states.map(function(s) {
        return '<button class="rbtn' + (gAccFilterState === s.key ? ' act' : '') + '"'
          + ' onclick="setAccFilterState(\'' + s.key + '\')"'
          + ' style="font-size:11px;padding:4px 14px">' + s.label + '</button>';
      }).join('');
}

function renderAccDateFilters() {
  var el = document.getElementById('acc-date-filters');
  if (!el) return;
  var dates = [
    { key: 'all',        label: 'All dates' },
    { key: 'this-month', label: 'This month' },
    { key: 'last-month', label: 'Last month' },
    { key: 'this-year',  label: 'This year' }
  ];
  el.innerHTML = '<span style="font-size:12px;font-weight:600;color:var(--g400);text-transform:uppercase;letter-spacing:.06em">Date:</span>'
    + dates.map(function(d) {
        return '<button class="rbtn' + (gAccFilterDate === d.key ? ' act' : '') + '"'
          + ' onclick="setAccFilterDate(\'' + d.key + '\')"'
          + ' style="font-size:11px;padding:4px 14px">' + d.label + '</button>';
      }).join('');
}

function accDateOk(dateStr) {
  if (gAccFilterDate === 'all' || !dateStr) return true;
  var now = new Date();
  var d   = new Date(dateStr);
  if (gAccFilterDate === 'this-month') return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
  if (gAccFilterDate === 'last-month') {
    var lm = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return d.getFullYear() === lm.getFullYear() && d.getMonth() === lm.getMonth();
  }
  if (gAccFilterDate === 'this-year') return d.getFullYear() === now.getFullYear();
  return true;
}

// ── Accounts Manager filter setters ──────────────────────
function setAccFilter(f) {
  gAccFilter = f;
  renderAccTypeFilters();
  renderAccMgrTable();
}

function setAccFilterState(s) {
  gAccFilterState = s;
  renderAccStateFilters();
  renderAccMgrTable();
}

function setAccFilterDate(d) {
  gAccFilterDate = d;
  renderAccDateFilters();
  renderAccMgrTable();
}

function setAccFilterPerson(email) {
  gAccFilterPerson = email;
  renderAccPersonFilters();
  renderAccMgrTable();
}

function toggleAllExpenses() {
  gShowAllExp = !gShowAllExp;
  renderAccMgrTable();
}

// ── Render: Accounts Manager card list ───────────────────
function renderAccMgrTable() {
  if (!gIsAccMgr) return;

  renderAccTypeFilters();
  renderAccStateFilters();
  renderAccDateFilters();
  renderAccPersonFilters();

  var all = [];
  gExpenses.forEach(function(e) { all.push({ type: 'standard', rec: e, date: e.date }); });
  gMileage.forEach(function(m)  { all.push({ type: 'mileage',  rec: m, date: m.date }); });
  all.sort(function(a, b) { return b.date.localeCompare(a.date); });

  var totalStd    = gExpenses.length;
  var totalMil    = gMileage.length;
  var unprocessed = gExpenses.filter(function(e) { return !e.processed; }).length;
  var ov = document.getElementById('acc-overview');
  if (ov) ov.innerHTML =
    '<span class="badge bgy" style="margin-right:6px">' + totalStd + ' standard</span>'
    + '<span class="badge bbl" style="margin-right:6px">' + totalMil + ' mileage</span>'
    + '<span class="badge bor">' + unprocessed + ' unprocessed</span>';

  var filtered = all.filter(function(item) {
    var typeOk =
      gAccFilter === 'all'
      || (gAccFilter === 'standard' && item.type === 'standard')
      || (gAccFilter === 'mileage'  && item.type === 'mileage');
    var stateOk =
      gAccFilterState === 'all'
      || (gAccFilterState === 'unprocessed' && !item.rec.processed)
      || (gAccFilterState === 'processed'   && item.rec.processed);
    var dateOk   = accDateOk(item.date);
    var personOk = !gAccFilterPerson || item.rec.person === gAccFilterPerson;
    return typeOk && stateOk && dateOk && personOk;
  });

  var el  = document.getElementById('acc-list');
  var saw = document.getElementById('acc-show-all-wrap');

  if (!filtered.length) {
    el.innerHTML = '<div class="empty">No expenses found.</div>';
    if (saw) saw.style.display = 'none';
    return;
  }

  var limit = gShowAllExp ? filtered.length : Math.min(10, filtered.length);
  var receiptSvg = '<svg width="18" height="18" viewBox="0 0 280 378" fill="currentColor" xmlns="http://www.w3.org/2000/svg" style="fill-rule:evenodd;clip-rule:evenodd;"><path d="M227.8,80.666l-176.171,0c-3.581,0 -6.51,2.93 -6.51,6.51c0,3.581 2.93,6.51 6.51,6.51l176.171,0c3.581,0 6.51,-2.93 6.51,-6.51c0,-3.581 -2.93,-6.51 -6.51,-6.51Z" style="fill-rule:nonzero;"/><path d="M227.8,151.758l-176.171,0c-3.581,0 -6.51,2.93 -6.51,6.51c0,3.581 2.93,6.51 6.51,6.51l176.171,0c3.581,0 6.51,-2.93 6.51,-6.51c0,-3.581 -2.93,-6.51 -6.51,-6.51Z" style="fill-rule:nonzero;"/><path d="M227.8,222.854l-176.171,0c-3.581,0 -6.51,2.93 -6.51,6.51c0,3.581 2.93,6.51 6.51,6.51l176.171,0c3.581,0 6.51,-2.93 6.51,-6.51c0,-3.581 -2.93,-6.51 -6.51,-6.51Z" style="fill-rule:nonzero;"/><path d="M233.008,0.002l-186.525,0c-25.651,0 -46.483,20.833 -46.483,46.483l0,318.229c0,2.604 1.562,4.948 3.971,5.99c2.409,1.042 5.143,0.521 7.031,-1.302l35.417,-34.05l42.121,40.43c2.539,2.409 6.51,2.409 8.984,0l42.121,-40.43l42.121,40.43c1.237,1.237 2.865,1.823 4.492,1.823c1.627,0 3.255,-0.586 4.492,-1.823l42.121,-40.43l35.547,34.115c1.888,1.823 4.688,2.344 7.031,1.302c2.409,-1.042 3.971,-3.385 3.971,-5.99l0,-318.296c0,-25.651 -20.833,-46.483 -46.483,-46.483l0.072,0.002Zm33.463,349.479l-29.037,-27.865c-2.539,-2.409 -6.51,-2.409 -8.984,0l-42.121,40.43l-42.121,-40.43c-2.539,-2.409 -6.51,-2.409 -8.984,0l-42.121,40.43l-42.121,-40.43c-1.237,-1.237 -2.865,-1.823 -4.492,-1.823c-1.627,0 -3.255,0.586 -4.492,1.823l-28.906,27.8l0,-302.929c0,-18.425 15.039,-33.463 33.463,-33.463l186.525,0c18.425,0 33.463,15.039 33.463,33.463l0,302.996l-0.073,-0.002Z" style="fill-rule:nonzero;"/></svg>';
  var carSvg = '<svg width="18" height="18" viewBox="0 0 200 176" fill="currentColor" xmlns="http://www.w3.org/2000/svg" style="fill-rule:evenodd;clip-rule:evenodd;"><path d="M4.167,142.936l3.902,0l0,28.719c0,2.301 1.864,4.167 4.167,4.167l26.424,0c2.303,0 4.167,-1.866 4.167,-4.167l0,-28.719l114.343,0l0,28.719c0,2.301 1.864,4.167 4.167,4.167l26.424,0c2.303,0 4.167,-1.866 4.167,-4.167l0,-28.719l3.906,0c2.303,0 4.167,-1.866 4.167,-4.167l0,-50.653c0,-6.56 -4.684,-12.042 -10.883,-13.289l-18.573,-64.282c-1.79,-6.209 -7.556,-10.545 -14.018,-10.545l-113.883,0c-6.462,0 -12.227,4.336 -14.018,10.543l-18.651,64.556c-5.73,1.584 -9.974,6.79 -9.974,13.018l0,50.653c0,2.301 1.864,4.167 4.167,4.167Zm30.326,24.552l-18.091,0l0,-24.552l18.091,0l0,24.552Zm149.101,0l-18.091,0l0,-24.552l18.091,0l0,24.552Zm-146.96,-154.637c0.765,-2.661 3.239,-4.519 6.01,-4.519l113.883,0c2.771,0 5.245,1.858 6.01,4.521l17.826,61.702l-161.556,0l17.826,-61.705Zm-28.3,75.264c0,-2.883 2.344,-5.227 5.225,-5.227l172.88,0c2.885,0 5.229,2.344 5.229,5.227l0,46.486l-183.333,0l-0,-46.486Z" style="fill-rule:nonzero;"/><path d="M31.954,122.19c7.414,0 13.444,-6.03 13.444,-13.444c0,-7.412 -6.03,-13.442 -13.444,-13.442c-7.41,0 -13.44,6.03 -13.44,13.442c0,7.414 6.03,13.444 13.44,13.444Zm0,-18.553c2.82,0 5.111,2.291 5.111,5.109c0,2.818 -2.291,5.111 -5.111,5.111c-2.816,0 -5.107,-2.293 -5.107,-5.111c0,-2.818 2.291,-5.109 5.107,-5.109Z" style="fill-rule:nonzero;"/><path d="M168.042,122.19c7.41,0 13.44,-6.03 13.44,-13.444c0,-7.412 -6.03,-13.442 -13.44,-13.442c-7.414,0 -13.444,6.03 -13.444,13.442c0,7.414 6.03,13.444 13.444,13.444Zm0,-18.553c2.816,0 5.107,2.291 5.107,5.109c0,2.818 -2.291,5.111 -5.107,5.111c-2.82,0 -5.111,-2.293 -5.111,-5.111c0,-2.818 2.291,-5.109 5.111,-5.109Z" style="fill-rule:nonzero;"/><path d="M71.533,122.19c2.303,0 4.167,-1.866 4.167,-4.167l0,-18.553c0,-2.301 -1.864,-4.167 -4.167,-4.167c-2.303,0 -4.167,1.866 -4.167,4.167l0,18.553c0,2.301 1.864,4.167 4.167,4.167Z" style="fill-rule:nonzero;"/><path d="M90.511,122.19c2.303,0 4.167,-1.866 4.167,-4.167l0,-18.553c0,-2.301 -1.864,-4.167 -4.167,-4.167c-2.303,0 -4.167,1.866 -4.167,4.167l0,18.553c0,2.301 1.864,4.167 4.167,4.167Z" style="fill-rule:nonzero;"/><path d="M109.485,122.19c2.303,0 4.167,-1.866 4.167,-4.167l0,-18.553c0,-2.301 -1.864,-4.167 -4.167,-4.167c-2.303,0 -4.167,1.866 -4.167,4.167l0,18.553c-0,2.301 1.864,4.167 4.167,4.167Z" style="fill-rule:nonzero;"/><path d="M128.463,122.19c2.303,0 4.167,-1.866 4.167,-4.167l0,-18.553c0,-2.301 -1.864,-4.167 -4.167,-4.167c-2.303,0 -4.167,1.866 -4.167,4.167l0,18.553c0,2.301 1.864,4.167 4.167,4.167Z" style="fill-rule:nonzero;"/></svg>';
  el.innerHTML = filtered.slice(0, limit).map(function(item) {
    var r          = item.rec;
    var personName = escHtml(cleanName(nameByEmail(r.person) || r.person));
    var cp         = escHtml(r.client) + '-' + escHtml(r.project);

    if (item.type === 'standard') {
      var statusBadge = r.processed
        ? '<span class="badge bgr" style="font-size:10px">Processed</span>'
        : '<span class="badge bor" style="font-size:10px">Pending</span>';
      var markBtn = !r.processed
        ? '<button class="rbtn" style="font-size:10px;padding:2px 8px"'
            + ' onclick="accMarkProcessed(\'standard\',\'' + escHtml(r.id) + '\')">&#10003; Processed</button>'
        : '';
      var dlBtn = r.hasAttachment
        ? '<button class="rbtn" style="font-size:10px;padding:2px 8px"'
            + ' onclick="accDownloadAttachment(\'' + escHtml(r.id) + '\')">&#8595; Receipt</button>'
        : '<button class="rbtn" style="font-size:10px;padding:2px 8px;opacity:.35;cursor:not-allowed" disabled>&#8595; Receipt</button>';
      var exVat = r.incVat ? r.amountExVat.toFixed(2) : r.amount.toFixed(2);
      var vat   = r.incVat ? r.vatRate : 0;
      return '<div class="tc-compact">'
        + '<div class="acc-icon" style="color:var(--btx)">' + receiptSvg + '</div>'
        + '<div style="flex:1;min-width:0">'
          + '<div style="display:flex;align-items:baseline;justify-content:space-between;gap:8px">'
            + '<div style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px">'
              + '<span style="font-weight:500;color:var(--g900)">' + personName + '</span>'
              + '<span style="color:var(--g400)"> · ' + cp + ' · ' + fmt(r.date) + ' · </span>'
              + '<span style="color:var(--g600)">' + escHtml(r.title) + '</span>'
            + '</div>'
            + '<div style="font-size:13px;font-weight:600;color:var(--or);white-space:nowrap;flex-shrink:0">£' + r.amount.toFixed(2) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;justify-content:space-between;gap:4px;margin-top:4px;flex-wrap:wrap">'
            + '<div style="display:flex;gap:4px;align-items:center;flex-wrap:wrap">'
              + statusBadge + markBtn + dlBtn
            + '</div>'
            + '<span style="font-size:10.5px;color:var(--g400)">Ex-VAT £' + exVat + ' · VAT ' + vat + '%</span>'
          + '</div>'
        + '</div>'
        + '</div>';
    } else {
      var tot = r.distance * r.rate;
      var mStatusBadge = r.processed
        ? '<span class="badge bgr" style="font-size:10px">Processed</span>'
        : '<span class="badge bor" style="font-size:10px">Pending</span>';
      var mMarkBtn = !r.processed
        ? '<button class="rbtn" style="font-size:10px;padding:2px 8px"'
            + ' onclick="accMarkProcessed(\'mileage\',\'' + escHtml(r.id) + '\')">&#10003; Processed</button>'
        : '';
      var mDlBtn = '<button class="rbtn" style="font-size:10px;padding:2px 8px;opacity:.35;cursor:not-allowed" disabled>&#8595; Receipt</button>';
      return '<div class="tc-compact">'
        + '<div class="acc-icon" style="color:var(--g600)">' + carSvg + '</div>'
        + '<div style="flex:1;min-width:0">'
          + '<div style="display:flex;align-items:baseline;justify-content:space-between;gap:8px">'
            + '<div style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px">'
              + '<span style="font-weight:500;color:var(--g900)">' + personName + '</span>'
              + '<span style="color:var(--g400)"> · ' + cp + ' · ' + fmt(r.date) + ' · ' + r.distance + 'mi · </span>'
              + '<span style="color:var(--g600)">' + escHtml(r.title) + '</span>'
            + '</div>'
            + '<div style="font-size:13px;font-weight:600;color:var(--or);white-space:nowrap;flex-shrink:0">£' + tot.toFixed(2) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;justify-content:space-between;gap:4px;margin-top:4px;flex-wrap:wrap">'
            + '<div style="display:flex;gap:4px;align-items:center;flex-wrap:wrap">'
              + mStatusBadge + mMarkBtn + mDlBtn
            + '</div>'
            + '<span style="font-size:10.5px;color:var(--g400)">' + r.distance + ' mi · £' + r.rate.toFixed(4) + '/mi</span>'
          + '</div>'
        + '</div>'
        + '</div>';
    }
  }).join('');

  if (saw) {
    saw.style.display = filtered.length > 10 ? '' : 'none';
    var sawBtn = saw.querySelector('button');
    if (sawBtn) sawBtn.textContent = gShowAllExp ? 'Hide expenses' : 'Show all expenses';
  }
}

// ── Accounts Manager actions ──────────────────────────────
function accMarkProcessed(type, id) {
  var listId = type === 'standard' ? gELId : gMLId;
  gPatch('/sites/' + gSiteId + '/lists/' + listId + '/items/' + id + '/fields', { is_processed: true })
    .then(function() {
      var arr = type === 'standard' ? gExpenses : gMileage;
      for (var i = 0; i < arr.length; i++) {
        if (arr[i].id === id) { arr[i].processed = true; break; }
      }
      renderAccMgrTable();
      expShowToast('Marked as processed.');
    })
    .catch(function(e) { expShowToast('Error: ' + e.message, true); });
}

function accDownloadAttachment(id) {
  var siteHost = CFG.tenant + '.sharepoint.com';
  var spScope  = 'https://' + siteHost + '/.default';
  var acct     = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  gMsal.acquireTokenSilent({ scopes: [spScope], account: acct }).then(function(spRes) {
    var tok      = spRes.accessToken;
    var sitePath = CFG.siteUrl.replace('https://' + siteHost, '');
    var url      = 'https://' + siteHost + sitePath
      + '/_api/web/lists/getbytitle(\'' + CFG.listExpenses + '\')/items(' + id + ')/AttachmentFiles';
    return fetch(url, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=verbose' }
    }).then(function(r) { return r.json(); }).then(function(data) {
      var files = (data.d && data.d.results) || [];
      if (!files.length) { expShowToast('No attachment found.', true); return; }
      window.open('https://' + siteHost + files[0].ServerRelativeUrl, '_blank', 'noopener,noreferrer');
    });
  }).catch(function(e) { expShowToast('Could not download attachment: ' + e.message, true); });
}
