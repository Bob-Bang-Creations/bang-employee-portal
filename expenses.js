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
        processed:   !!f.is_processed,
        person:      (f.person || '').toLowerCase()
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
        distance: parseFloat(f.distance) || 0,
        rate:     parseFloat(f.rate) || 0.45,
        person:   (f.person || '').toLowerCase()
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

  // If rate typed as 0 while checkbox is ticked, auto-untick
  if (rate === 0 && cbEl.checked) cbEl.checked = false;
  // If rate > 0 while checkbox is unticked, auto-tick
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

// ── Add standard expense ─────────────────────────────────
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
  if (!receipt) { alert('Please attach a receipt (PDF or image).'); return; }

  var exVat  = hasVat ? amt / (1 + vatRate / 100) : amt;
  var vatAmt = hasVat ? amt - exVat : 0;

  var btn = document.getElementById('es-addbtn');
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

  // Step 1: create the list item
  gPost('/sites/' + gSiteId + '/lists/' + gELId + '/items', {
    fields: {
      Title:          title,
      expense_date:   date,
      client_number:  client,
      project_number: project,
      amount:         amt.toFixed(2),
      inc_VAT:        hasVat,
      VAT_rate:       hasVat ? vatRate.toFixed(2) : '0',
      amount_exVAT:   exVat.toFixed(2),
      amount_VAT:     vatAmt.toFixed(2),
      is_processed:   false,
      person:         gUser.email
    }
  }).then(function(r) {
    // Step 2: upload attachment via SharePoint REST
    // Requires a SharePoint-scoped token — Graph does not support list item attachments
    var siteHost = CFG.tenant + '.sharepoint.com';
    var spScope  = 'https://' + siteHost + '/.default';
    var acct     = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
    return gMsal.acquireTokenSilent({ scopes: [spScope], account: acct }).then(function(spRes) {
      var tok       = spRes.accessToken;
      var sitePath  = CFG.siteUrl.replace('https://' + siteHost, '');
      var attachUrl = 'https://' + siteHost + sitePath
        + '/_api/web/lists/getbytitle(\'' + CFG.listExpenses + '\')/items(' + r.id + ')/AttachmentFiles/add(FileName=\'' + encodeURIComponent(receipt.name) + '\')';
      return new Promise(function(resolve, reject) {
        var reader = new FileReader();
        reader.onload = function(e) {
          fetch(attachUrl, {
            method:  'POST',
            headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/octet-stream' },
            body:    e.target.result
          }).then(function(res) {
            if (!res.ok) return res.text().then(function(t) { reject(new Error('Attachment upload failed: ' + t)); });
            resolve(r);
          }).catch(reject);
        };
        reader.onerror = function() { reject(new Error('Failed to read receipt file.')); };
        reader.readAsArrayBuffer(receipt);
      });
    }).then(function() { return r; });
  }).then(function(r) {
    gExpenses.push({
      id: r.id, title: title, date: date, client: client, project: project,
      amount: amt, amountExVat: exVat, amountVat: vatAmt,
      incVat: hasVat, vatRate: vatRate, processed: false, person: gUser.email
    });
    // Reset form
    document.getElementById('es-client').value         = '';
    document.getElementById('es-project').value        = '';
    document.getElementById('es-date').value           = '';
    document.getElementById('es-title').value          = '';
    document.getElementById('es-amount').value         = '';
    document.getElementById('es-hasvat').checked       = true;
    document.getElementById('es-vatrate').value        = '20';
    document.getElementById('es-vat-preview').textContent = '';
    document.getElementById('es-receipt').value        = '';
    renderMyExpenses();
    renderAccMgrTable();
    expShowToast('Expense logged successfully!');
  }).catch(function(e) {
    expShowToast('Error: ' + e.message, true);
  }).then(function() {
    btn.disabled = false; btn.textContent = 'Log expense';
  });
}

// ── Add mileage expense ──────────────────────────────────
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
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

  gPost('/sites/' + gSiteId + '/lists/' + gMLId + '/items', {
    fields: {
      Title:          title,
      travel_date:    date,
      client_number:  client,
      project_number: project,
      distance:       dist.toString(),
      rate:           rate.toFixed(4),
      person:         gUser.email
    }
  }).then(function(r) {
    gMileage.push({ id: r.id, title: title, date: date, client: client, project: project, distance: dist, rate: rate, person: gUser.email });
    document.getElementById('em-client').value       = '';
    document.getElementById('em-project').value      = '';
    document.getElementById('em-date').value         = '';
    document.getElementById('em-title').value        = '';
    document.getElementById('em-distance').value     = '';
    document.getElementById('em-rate').value         = '0.45';
    document.getElementById('em-preview').textContent = '';
    renderMyExpenses();
    renderAccMgrTable();
    expShowToast('Mileage logged successfully!');
  }).catch(function(e) {
    expShowToast('Error: ' + e.message, true);
  }).then(function() {
    btn.disabled = false; btn.textContent = 'Log mileage';
  });
}

// ── Render: My expenses ──────────────────────────────────
function renderMyExpenses() {
  var email  = gUser ? gUser.email : '';
  var myStd  = gExpenses.filter(function(e) { return e.person === email; });
  var myMil  = gMileage.filter(function(m)  { return m.person === email; });

  var all = [];
  myStd.forEach(function(e) { all.push({ type: 'standard', rec: e, date: e.date }); });
  myMil.forEach(function(m) { all.push({ type: 'mileage',  rec: m, date: m.date }); });
  all.sort(function(a, b) { return b.date.localeCompare(a.date); });

  var el = document.getElementById('my-expenses-list');
  if (!all.length) { el.innerHTML = '<div class="empty">No expenses logged yet.</div>'; return; }

  el.innerHTML = all.map(function(item) {
    var r = item.rec;
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
        + '<div style="font-size:16px;font-weight:600;color:var(--or);white-space:nowrap">' + total + '</div>'
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
        + '<div style="font-size:16px;font-weight:600;color:var(--or);white-space:nowrap">£' + tot.toFixed(2) + '</div>'
        + '</div>'
        + '</div>';
    }
  }).join('');
}

// ── Render: Accounts Manager table ──────────────────────
function setAccFilter(f) {
  gAccFilter = f;
  renderAccMgrTable();
}

function toggleAllExpenses() {
  gShowAllExp = !gShowAllExp;
  renderAccMgrTable();
}

function renderAccMgrTable() {
  if (!gIsAccMgr) return;

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
    if (gAccFilter === 'standard')    return item.type === 'standard';
    if (gAccFilter === 'mileage')     return item.type === 'mileage';
    if (gAccFilter === 'unprocessed') return item.type === 'standard' && !item.rec.processed;
    return true;
  });

  var tb  = document.getElementById('acc-tbody');
  var saw = document.getElementById('acc-show-all-wrap');

  if (!filtered.length) {
    tb.innerHTML = '<tr><td colspan="10" class="empty">No expenses found.</td></tr>';
    if (saw) saw.style.display = 'none';
    return;
  }

  var limit = gShowAllExp ? filtered.length : Math.min(10, filtered.length);
  tb.innerHTML = filtered.slice(0, limit).map(function(item) {
    var r          = item.rec;
    var personName = escHtml(nameByEmail(r.person) || r.person);

    if (item.type === 'standard') {
      var status = r.processed
        ? '<span class="badge bgr" style="font-size:10px">Processed</span>'
        : '<span class="badge bor" style="font-size:10px">Pending</span>';
      var exVat = r.incVat ? '£' + r.amountExVat.toFixed(2) : '—';
      var vat   = r.incVat ? '£' + r.amountVat.toFixed(2)   : '—';
      return '<tr>'
        + '<td title="' + personName + '">' + personName + '</td>'
        + '<td title="' + escHtml(r.title) + '">' + escHtml(r.title) + '</td>'
        + '<td>' + escHtml(r.client) + '</td>'
        + '<td>' + escHtml(r.project) + '</td>'
        + '<td>' + fmt(r.date) + '</td>'
        + '<td><span class="badge bbl" style="font-size:10px">Standard</span></td>'
        + '<td style="font-weight:600;color:var(--or)">£' + r.amount.toFixed(2) + '</td>'
        + '<td>' + exVat + '</td>'
        + '<td>' + vat + '</td>'
        + '<td>' + status + '</td>'
        + '</tr>';
    } else {
      var tot = r.distance * r.rate;
      return '<tr>'
        + '<td title="' + personName + '">' + personName + '</td>'
        + '<td title="' + escHtml(r.title) + '">' + escHtml(r.title) + ' (' + r.distance + 'mi)</td>'
        + '<td>' + escHtml(r.client) + '</td>'
        + '<td>' + escHtml(r.project) + '</td>'
        + '<td>' + fmt(r.date) + '</td>'
        + '<td><span class="badge bgy" style="font-size:10px">Mileage</span></td>'
        + '<td style="font-weight:600;color:var(--or)">£' + tot.toFixed(2) + '</td>'
        + '<td>—</td>'
        + '<td>—</td>'
        + '<td><span class="badge bgr" style="font-size:10px">N/A</span></td>'
        + '</tr>';
    }
  }).join('');

  if (saw) {
    saw.style.display = filtered.length > 10 ? '' : 'none';
    var sawBtn = saw.querySelector('button');
    if (sawBtn) sawBtn.textContent = gShowAllExp ? 'Hide expenses' : 'Show all expenses';
  }
}
