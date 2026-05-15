// ── HR Tasks ─────────────────────────────────────────────
// Depends on: config.js, graph.js

// ── Fetch ────────────────────────────────────────────────
function fetchTasks() {
  return gGet('/sites/' + gSiteId + '/lists/' + gTLId + '/items?$expand=fields&$top=500').then(function(d) {
    gTasks = d.value.map(function(i) {
      return {
        id:            i.id,
        title:         i.fields.TaskTitle || i.fields.Title || '',
        desc:          i.fields.Description || '',
        due:           (i.fields.DueDate || '').split('T')[0],
        cat:           i.fields.Category || '',
        assignees:     tryParse(i.fields.Assignees, []),
        links:         tryParse(i.fields.Links, []),
        hasAttachment: !!i.fields.Attachments,
        completed:     !!i.fields.completed,
        completedOn:   i.fields.CompletedOn || null
      };
    });
  });
}

// ── Helpers ──────────────────────────────────────────────
function isOver(t) {
  return t.due < today() && !t.completed;
}

function dueBadge(t) {
  if (t.completed)     return '<span class="badge bgr">&#10003; Complete</span>';
  if (t.due < today()) return '<span class="badge bre">Overdue</span>';
  var d = Math.ceil((new Date(t.due) - new Date(today())) / 86400000);
  if (d <= 7)          return '<span class="badge bor">Due soon</span>';
  return '<span class="badge bgy">Upcoming</span>';
}

function sc(n, l, key) {
  var act = gMyFilter === key ? ' act' : '';
  return '<div class="sc' + act + '" onclick="setMyFilter(\'' + key + '\')">'
    + '<div class="n">' + n + '</div>'
    + '<div class="l">' + l + '</div>'
    + '</div>';
}

function setMyFilter(key) {
  gMyFilter = gMyFilter === key ? null : key;
  renderMy();
}

// ── Render: My Tasks ─────────────────────────────────────
function renderMy() {
  var email = gUser ? gUser.email : '';
  var mt    = gTasks.filter(function(t) {
    return t.assignees.some(function(a) { return a.toLowerCase() === email; });
  });
  var dn = mt.filter(function(t) { return t.completed; }).length;
  var ov = mt.filter(function(t) { return isOver(t); }).length;

  document.getElementById('my-stats').innerHTML =
    sc(mt.length - dn, 'Pending', 'pending') + sc(ov, 'Overdue', 'overdue') + sc(dn, 'Complete', 'complete');

  var displayed = mt;
  if (gMyFilter === 'pending')  displayed = mt.filter(function(t) { return !t.completed; });
  if (gMyFilter === 'overdue')  displayed = mt.filter(function(t) { return isOver(t); });
  if (gMyFilter === 'complete') displayed = mt.filter(function(t) { return t.completed; });

  if (!displayed.length) {
    document.getElementById('my-list').innerHTML = '<div class="empty">' + (mt.length ? 'No tasks match this filter.' : 'No tasks assigned to you yet.') + '</div>';
    return;
  }

  document.getElementById('my-list').innerHTML = displayed.map(function(t) {
    var d  = t.completed;
    var lk = t.links.length
      ? '<div class="tlinks">' + t.links.map(function(l) {
          var u = safeUrl(l);
          return '<a href="' + escHtml(u) + '" target="_blank" rel="noopener noreferrer">&#128279; ' + escHtml(l.replace(/^https?:\/\//, '').split('/')[0]) + '</a>';
        }).join('') + '</div>'
      : '';
    var dlBtn = t.hasAttachment
      ? '<div style="margin-top:6px"><button class="rbtn" style="font-size:10px;padding:2px 10px" onclick="taskDownloadAttachment(\'' + escHtml(t.id) + '\')">&#8595; Download task attachment</button></div>'
      : '';
    return '<div class="tc' + (d ? ' done' : '') + '">'
      + '<div class="ttop">'
      + '<div style="flex:1;min-width:0">'
      + '<div class="tn' + (d ? ' struck' : '') + '">' + escHtml(t.title) + '</div>'
      + '<div class="tmeta">'
      + dueBadge(t)
      + '<span class="badge bbl">' + escHtml(t.cat) + '</span>'
      + '<span style="font-size:11.5px;color:var(--g400)">Due ' + fmt(t.due) + '</span>'
      + '</div>'
      + '<div class="tdesc">' + escHtml(t.desc) + '</div>'
      + lk
      + dlBtn
      + '</div>'
      + '<button class="cbtn' + (d ? ' dn' : '') + '" id="cb-' + t.id + '" data-tid="' + t.id + '" onclick="toggleComp(this.dataset.tid)">' + (d ? '&#10003; Done' : 'Mark done') + '</button>'
      + '</div>'
      + '</div>';
  }).join('');
}

// ── Render: Manage table ─────────────────────────────────
function renderManage() {
  var tb  = document.getElementById('mtbody');
  var ov  = document.getElementById('manage-overview');
  var saw = document.getElementById('show-all-wrap');

  if (!gTasks.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">No tasks yet.</td></tr>';
    if (ov)  ov.innerHTML = '';
    if (saw) saw.style.display = 'none';
    return;
  }

  var incomplete = gTasks.filter(function(t) { return !t.completed; }).length;
  var overdue    = gTasks.filter(function(t) { return isOver(t); }).length;

  if (ov) ov.innerHTML =
    '<span class="badge bgy" style="margin-right:6px">' + incomplete + ' incomplete</span>'
    + '<span class="badge bre">' + overdue + ' overdue</span>';

  var rows = gTasks.map(function(t) {
    var names = t.assignees.map(function(e) {
      var n  = cleanName(nameByEmail(e));
      var cl = t.completed ? 'bgr' : isOver(t) ? 'bre' : 'bgy';
      return '<span class="badge ' + cl + '" style="font-size:10px">' + escHtml(n) + '</span>';
    }).join(' ');
    return '<tr>'
      + '<td title="' + escHtml(t.title) + '" style="font-weight:500">' + escHtml(t.title) + '</td>'
      + '<td><div style="display:flex;flex-wrap:wrap;gap:3px">' + names + '</div></td>'
      + '<td>' + fmt(t.due) + '</td>'
      + '<td><span class="badge bbl" style="font-size:10px">' + escHtml(t.cat) + '</span></td>'
      + '<td style="color:var(--g600)">' + (t.completed ? 'Done' : 'Pending') + '</td>'
      + '<td style="white-space:nowrap">'
        + '<button class="rbtn" style="font-size:10px;padding:2px 8px;margin-right:4px" onclick="taskEdit(\'' + escHtml(t.id) + '\')">Edit</button>'
        + '<button class="dbtn" id="db-' + t.id + '" data-tid="' + t.id + '" onclick="delTask(this.dataset.tid)">Delete</button>'
      + '</td>'
      + '</tr>';
  });

  var limit = gShowAllTasks ? rows.length : Math.min(5, rows.length);
  tb.innerHTML = rows.slice(0, limit).join('');
  if (saw) {
    saw.style.display = gTasks.length > 5 ? '' : 'none';
    var sawBtn = saw.querySelector('button');
    if (sawBtn) sawBtn.textContent = gShowAllTasks ? 'Hide tasks' : 'Show all tasks';
  }
}

// ── Assign form ──────────────────────────────────────────
function toggleAssignForm() {
  var s       = document.getElementById('assign-form-section');
  var opening = s.style.display === 'none';
  s.style.display = opening ? '' : 'none';
  var btn = document.getElementById('assign-toggle-btn');
  if (btn) btn.textContent = opening ? '− Assign new training task' : '+ Assign new training task';
  if (opening) {
    var srch = document.getElementById('assignee-search');
    if (srch) { srch.value = ''; filterAssignees(''); }
  } else {
    gEditingTaskId = null;
    document.getElementById('addbtn').textContent = 'Assign task';
  }
}

function resetTaskForm() {
  document.getElementById('f-title').value      = '';
  document.getElementById('f-desc').value       = '';
  document.getElementById('f-due').value        = '';
  document.getElementById('f-links').value      = '';
  document.getElementById('f-attachment').value = '';
  clearAllAssignees();
  document.getElementById('assign-form-section').style.display = 'none';
  var btn = document.getElementById('assign-toggle-btn');
  if (btn) btn.textContent = '+ Assign new training task';
}

function taskEdit(tid) {
  var t = null;
  for (var i = 0; i < gTasks.length; i++) { if (gTasks[i].id === tid) { t = gTasks[i]; break; } }
  if (!t) return;
  document.getElementById('f-title').value = t.title;
  document.getElementById('f-desc').value  = t.desc;
  document.getElementById('f-due').value   = t.due;
  document.getElementById('f-links').value = t.links.join(', ');
  document.getElementById('f-attachment').value = '';
  var cbs = document.querySelectorAll('#assignee-list input[type=checkbox]');
  for (var j = 0; j < cbs.length; j++) {
    var chk = t.assignees.indexOf(cbs[j].value) >= 0;
    cbs[j].checked = chk;
    if (chk) cbs[j].closest('.assignee-item').classList.add('checked');
    else     cbs[j].closest('.assignee-item').classList.remove('checked');
  }
  gEditingTaskId = tid;
  document.getElementById('addbtn').textContent = 'Update task';
  var s = document.getElementById('assign-form-section');
  s.style.display = '';
  var toggleBtn = document.getElementById('assign-toggle-btn');
  if (toggleBtn) toggleBtn.textContent = '− Assign new training task';
  s.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function taskDownloadAttachment(id) {
  var siteHost = CFG.tenant + '.sharepoint.com';
  var spScope  = 'https://' + siteHost + '/.default';
  var acct     = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  gMsal.acquireTokenSilent({ scopes: [spScope], account: acct }).then(function(spRes) {
    var tok      = spRes.accessToken;
    var sitePath = CFG.siteUrl.replace('https://' + siteHost, '');
    var url      = 'https://' + siteHost + sitePath
      + '/_api/web/lists/getbytitle(\'' + CFG.listTasks + '\')/items(' + id + ')/AttachmentFiles';
    return fetch(url, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=verbose' }
    }).then(function(r) { return r.json(); }).then(function(data) {
      var files = (data.d && data.d.results) || [];
      if (!files.length) { showToast('No attachment found.', true); return; }
      window.open('https://' + siteHost + files[0].ServerRelativeUrl, '_blank', 'noopener,noreferrer');
    });
  }).catch(function(e) { showToast('Could not download attachment: ' + e.message, true); });
}

function uploadTaskAttachment(itemId, file) {
  var siteHost = CFG.tenant + '.sharepoint.com';
  var spScope  = 'https://' + siteHost + '/.default';
  var acct     = gMsal.getActiveAccount() || gMsal.getAllAccounts()[0];
  return gMsal.acquireTokenSilent({ scopes: [spScope], account: acct }).then(function(spRes) {
    var tok      = spRes.accessToken;
    var sitePath = CFG.siteUrl.replace('https://' + siteHost, '');
    var url      = 'https://' + siteHost + sitePath
      + '/_api/web/lists/getbytitle(\'' + CFG.listTasks + '\')/items(' + itemId + ')/AttachmentFiles/add(FileName=\'' + encodeURIComponent(file.name) + '\')';
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
      reader.onerror = function() { reject(new Error('Failed to read attachment file.')); };
      reader.readAsArrayBuffer(file);
    });
  });
}

function filterAssignees(q) {
  var term  = q.toLowerCase().trim();
  var items = document.querySelectorAll('#assignee-list .assignee-item');
  for (var i = 0; i < items.length; i++) {
    var name  = (items[i].querySelector('.aname')  || {}).textContent || '';
    var email = (items[i].querySelector('.aemail') || {}).textContent || '';
    items[i].style.display = (!term || name.toLowerCase().indexOf(term) >= 0 || email.toLowerCase().indexOf(term) >= 0) ? '' : 'none';
  }
}

function toggleAllTasks() {
  gShowAllTasks = !gShowAllTasks;
  renderManage();
}

function populateUserDropdown() {
  var list = document.getElementById('assignee-list');
  if (!gUsers.length) {
    list.innerHTML = '<div class="assignee-item" style="color:var(--g400);font-size:13px;justify-content:center;cursor:default">No users found</div>';
    return;
  }
  list.innerHTML = gUsers.map(function(u) {
    var safeEmail = escHtml(u.email);
    var safeName  = escHtml(u.name);
    return '<label class="assignee-item">'
      + '<input type="checkbox" value="' + safeEmail + '" onchange="onAssigneeChange(this)">'
      + '<div><div class="aname">' + safeName + '</div><div class="aemail">' + safeEmail + '</div></div>'
      + '</label>';
  }).join('');
  var hint = document.getElementById('user-load-hint');
  if (hint) hint.textContent = gUsers.length + ' users available';
}

function onAssigneeChange(cb) {
  var item = cb.closest('.assignee-item');
  if (cb.checked) item.classList.add('checked');
  else item.classList.remove('checked');
}

function selectAllAssignees() {
  var cbs = document.querySelectorAll('#assignee-list input[type=checkbox]');
  for (var i = 0; i < cbs.length; i++) {
    cbs[i].checked = true;
    cbs[i].closest('.assignee-item').classList.add('checked');
  }
}

function clearAllAssignees() {
  var cbs = document.querySelectorAll('#assignee-list input[type=checkbox]');
  for (var i = 0; i < cbs.length; i++) {
    cbs[i].checked = false;
    cbs[i].closest('.assignee-item').classList.remove('checked');
  }
}

// ── Actions ──────────────────────────────────────────────
function toggleComp(tid) {
  var btn = document.getElementById('cb-' + tid);
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="isp"></span>'; }

  var task = null;
  for (var i = 0; i < gTasks.length; i++) {
    if (gTasks[i].id === tid) { task = gTasks[i]; break; }
  }
  if (!task) return;

  var nowComplete = !task.completed;
  var fields = nowComplete
    ? { completed: true,  CompletedOn: new Date().toISOString() }
    : { completed: false, CompletedOn: null };

  gPatch('/sites/' + gSiteId + '/lists/' + gTLId + '/items/' + tid + '/fields', fields)
    .then(function() {
      task.completed   = nowComplete;
      task.completedOn = nowComplete ? fields.CompletedOn : null;
      renderMy(); renderManage();
    })
    .catch(function(e) {
      showToast('Error: ' + e.message, true);
      if (btn) { btn.disabled = false; btn.textContent = 'Mark done'; }
    });
}

function addTask() {
  var title = document.getElementById('f-title').value.trim();
  var desc  = document.getElementById('f-desc').value.trim();
  var due   = document.getElementById('f-due').value;
  var cat   = document.getElementById('f-cat').value;
  var raw   = document.getElementById('f-links').value.trim();
  var links = raw ? raw.split(',').map(function(l) { return l.trim(); }).filter(Boolean) : [];
  var file  = document.getElementById('f-attachment').files[0];

  var cbs  = document.querySelectorAll('#assignee-list input[type=checkbox]:checked');
  var asgn = [];
  for (var i = 0; i < cbs.length; i++) asgn.push(cbs[i].value);

  if (!title)       { alert('Please enter a task title.'); return; }
  if (!due)         { alert('Please select a due date.'); return; }
  if (!asgn.length) { alert('Please assign to at least one person.'); return; }

  var btn = document.getElementById('addbtn');
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

  // ── Edit mode ─────────────────────────────────────────
  if (gEditingTaskId) {
    var editId = gEditingTaskId;
    var p = gPatch(
      '/sites/' + gSiteId + '/lists/' + gTLId + '/items/' + editId + '/fields',
      { TaskTitle: title, Description: desc, DueDate: due, Category: cat,
        Assignees: JSON.stringify(asgn), Links: JSON.stringify(links) }
    );
    if (file) p = p.then(function() { return uploadTaskAttachment(editId, file); });
    p.then(function() {
      for (var i = 0; i < gTasks.length; i++) {
        if (gTasks[i].id === editId) {
          var t = gTasks[i];
          t.title = title; t.desc = desc; t.due = due; t.cat = cat;
          t.assignees = asgn; t.links = links;
          if (file) t.hasAttachment = true;
          break;
        }
      }
      gEditingTaskId = null;
      resetTaskForm();
      renderMy(); renderManage();
      showToast('Task updated!');
    }).catch(function(e) { showToast('Error: ' + e.message, true); })
    .then(function() { btn.disabled = false; btn.textContent = 'Assign task'; });
    return;
  }

  // ── Add mode ──────────────────────────────────────────
  gPost('/sites/' + gSiteId + '/lists/' + gTLId + '/items', {
    fields: {
      TaskTitle:   title,
      Description: desc,
      DueDate:     due,
      Category:    cat,
      Assignees:   JSON.stringify(asgn),
      Links:       JSON.stringify(links)
    }
  }).then(function(r) {
    var p = file ? uploadTaskAttachment(r.id, file).then(function() { return r; }) : Promise.resolve(r);
    return p;
  }).then(function(r) {
    gTasks.push({ id: r.id, title: title, desc: desc, due: due, cat: cat, assignees: asgn, links: links, hasAttachment: !!file, completed: false, completedOn: null });
    resetTaskForm();
    renderMy(); renderManage();
    showToast('Task assigned successfully!');
  }).catch(function(e) { showToast('Error: ' + e.message, true); })
  .then(function() { btn.disabled = false; btn.textContent = 'Assign task'; });
}

function delTask(tid) {
  if (!confirm('Delete this task?')) return;
  var btn = document.getElementById('db-' + tid);
  if (btn) btn.disabled = true;
  gDel('/sites/' + gSiteId + '/lists/' + gTLId + '/items/' + tid)
    .then(function() {
      gTasks = gTasks.filter(function(t) { return t.id !== tid; });
      renderMy(); renderManage();
      showToast('Task deleted.');
    })
    .catch(function(e) { showToast('Error: ' + e.message, true); if (btn) btn.disabled = false; });
}

function showToast(msg, isErr) {
  var el = document.getElementById('toast'), tx = document.getElementById('toasttxt');
  tx.textContent = msg; el.classList.toggle('err', !!isErr); el.classList.add('show');
  setTimeout(function() { el.classList.remove('show'); }, 4000);
}
