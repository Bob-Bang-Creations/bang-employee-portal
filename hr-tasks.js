// ── HR Tasks ─────────────────────────────────────────────
// Depends on: config.js, graph.js

// ── Fetch ────────────────────────────────────────────────
function fetchTasks() {
  return gGet('/sites/' + gSiteId + '/lists/' + gTLId + '/items?$expand=fields&$top=500').then(function(d) {
    gTasks = d.value.map(function(i) {
      return {
        id:        i.id,
        title:     i.fields.TaskTitle || i.fields.Title || '',
        desc:      i.fields.Description || '',
        due:       (i.fields.DueDate || '').split('T')[0],
        cat:       i.fields.Category || '',
        assignees: tryParse(i.fields.Assignees, []),
        links:     tryParse(i.fields.Links, [])
      };
    });
  });
}

function fetchComps() {
  return gGet('/sites/' + gSiteId + '/lists/' + gCLId + '/items?$expand=fields&$top=1000').then(function(d) {
    gComps = d.value.map(function(i) {
      return {
        id:       i.id,
        taskId:   i.fields.TaskId || '',
        assignee: (i.fields.AssigneeName || '').toLowerCase()
      };
    });
  });
}

// ── Helpers ──────────────────────────────────────────────
function isComp(tid, email) {
  var e = email.toLowerCase();
  return gComps.some(function(c) { return c.taskId === String(tid) && c.assignee === e; });
}

function isOver(t, email) {
  return t.due < today() && !isComp(t.id, email);
}

function dueBadge(t, email) {
  if (isComp(t.id, email)) return '<span class="badge bgr">&#10003; Complete</span>';
  if (t.due < today())     return '<span class="badge bre">Overdue</span>';
  var d = Math.ceil((new Date(t.due) - new Date(today())) / 86400000);
  if (d <= 7)              return '<span class="badge bor">Due soon</span>';
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
  var dn = mt.filter(function(t) { return isComp(t.id, email); }).length;
  var ov = mt.filter(function(t) { return isOver(t, email); }).length;

  document.getElementById('my-stats').innerHTML =
    sc(mt.length - dn, 'Pending', 'pending') + sc(ov, 'Overdue', 'overdue') + sc(dn, 'Complete', 'complete');

  var displayed = mt;
  if (gMyFilter === 'pending')  displayed = mt.filter(function(t) { return !isComp(t.id, email); });
  if (gMyFilter === 'overdue')  displayed = mt.filter(function(t) { return isOver(t, email); });
  if (gMyFilter === 'complete') displayed = mt.filter(function(t) { return isComp(t.id, email); });

  if (!displayed.length) {
    document.getElementById('my-list').innerHTML = '<div class="empty">' + (mt.length ? 'No tasks match this filter.' : 'No tasks assigned to you yet.') + '</div>';
    return;
  }

  document.getElementById('my-list').innerHTML = displayed.map(function(t) {
    var d  = isComp(t.id, email);
    var lk = t.links.length
      ? '<div class="tlinks">' + t.links.map(function(l) {
          var u = safeUrl(l);
          return '<a href="' + escHtml(u) + '" target="_blank" rel="noopener noreferrer">&#128279; ' + escHtml(l.replace(/^https?:\/\//, '').split('/')[0]) + '</a>';
        }).join('') + '</div>'
      : '';
    return '<div class="tc' + (d ? ' done' : '') + '">'
      + '<div class="ttop">'
      + '<div style="flex:1;min-width:0">'
      + '<div class="tn' + (d ? ' struck' : '') + '">' + escHtml(t.title) + '</div>'
      + '<div class="tmeta">'
      + dueBadge(t, email)
      + '<span class="badge bbl">' + escHtml(t.cat) + '</span>'
      + '<span style="font-size:11.5px;color:var(--g400)">Due ' + fmt(t.due) + '</span>'
      + '</div>'
      + '<div class="tdesc">' + escHtml(t.desc) + '</div>'
      + lk
      + '</div>'
      + '<button class="cbtn' + (d ? ' dn' : '') + '" id="cb-' + t.id + '" data-tid="' + t.id + '" data-email="' + escHtml(email) + '" onclick="toggleComp(this.dataset.tid,this.dataset.email)">' + (d ? '&#10003; Done' : 'Mark done') + '</button>'
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

  var incomplete = gTasks.filter(function(t) {
    return t.assignees.some(function(e) { return !isComp(t.id, e); });
  }).length;
  var overdue = gTasks.filter(function(t) {
    return t.due < today() && t.assignees.some(function(e) { return !isComp(t.id, e); });
  }).length;

  if (ov) ov.innerHTML =
    '<span class="badge bgy" style="margin-right:6px">' + incomplete + ' incomplete</span>'
    + '<span class="badge bre">' + overdue + ' overdue</span>';

  var rows = gTasks.map(function(t) {
    var dc    = t.assignees.filter(function(e) { return isComp(t.id, e); }).length;
    var names = t.assignees.map(function(e) {
      var n  = cleanName(nameByEmail(e));
      var cl = isComp(t.id, e) ? 'bgr' : isOver(t, e) ? 'bre' : 'bgy';
      return '<span class="badge ' + cl + '" style="font-size:10px">' + escHtml(n) + '</span>';
    }).join(' ');
    return '<tr>'
      + '<td title="' + escHtml(t.title) + '" style="font-weight:500">' + escHtml(t.title) + '</td>'
      + '<td><div style="display:flex;flex-wrap:wrap;gap:3px">' + names + '</div></td>'
      + '<td>' + fmt(t.due) + '</td>'
      + '<td><span class="badge bbl" style="font-size:10px">' + escHtml(t.cat) + '</span></td>'
      + '<td style="color:var(--g600)">' + dc + '/' + t.assignees.length + ' done</td>'
      + '<td><button class="dbtn" id="db-' + t.id + '" data-tid="' + t.id + '" onclick="delTask(this.dataset.tid)">Delete</button></td>'
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
  }
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
function toggleComp(tid, email) {
  var btn = document.getElementById('cb-' + tid);
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="isp"></span>'; }
  var ex = null;
  for (var i = 0; i < gComps.length; i++) {
    if (gComps[i].taskId === String(tid) && gComps[i].assignee === email.toLowerCase()) { ex = gComps[i]; break; }
  }
  var p = ex
    ? gDel('/sites/' + gSiteId + '/lists/' + gCLId + '/items/' + ex.id).then(function() {
        gComps = gComps.filter(function(c) { return c.id !== ex.id; });
      })
    : gPost('/sites/' + gSiteId + '/lists/' + gCLId + '/items', {
        fields: { TaskId: String(tid), AssigneeName: email.toLowerCase(), CompletedOn: new Date().toISOString() }
      }).then(function(r) { gComps.push({ id: r.id, taskId: String(tid), assignee: email.toLowerCase() }); });

  p.then(function() { renderMy(); renderManage(); })
   .catch(function(e) { showToast('Error: ' + e.message, true); if (btn) { btn.disabled = false; btn.textContent = 'Mark done'; } });
}

function addTask() {
  var title = document.getElementById('f-title').value.trim();
  var desc  = document.getElementById('f-desc').value.trim();
  var due   = document.getElementById('f-due').value;
  var cat   = document.getElementById('f-cat').value;
  var raw   = document.getElementById('f-links').value.trim();
  var links = raw ? raw.split(',').map(function(l) { return l.trim(); }).filter(Boolean) : [];

  var cbs  = document.querySelectorAll('#assignee-list input[type=checkbox]:checked');
  var asgn = [];
  for (var i = 0; i < cbs.length; i++) asgn.push(cbs[i].value);

  if (!title)       { alert('Please enter a task title.'); return; }
  if (!due)         { alert('Please select a due date.'); return; }
  if (!asgn.length) { alert('Please assign to at least one person.'); return; }

  var btn = document.getElementById('addbtn');
  btn.disabled = true; btn.innerHTML = '<span class="isp"></span> Saving…';

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
    gTasks.push({ id: r.id, title: title, desc: desc, due: due, cat: cat, assignees: asgn, links: links });
    document.getElementById('f-title').value = '';
    document.getElementById('f-desc').value  = '';
    document.getElementById('f-due').value   = '';
    document.getElementById('f-links').value = '';
    clearAllAssignees();
    document.getElementById('assign-form-section').style.display = 'none';
    renderMy(); renderManage();
    showToast('Task assigned successfully!');
  }).catch(function(e) { showToast('Error: ' + e.message, true); })
  .then(function() { btn.disabled = false; btn.textContent = 'Assign task'; });
}

function delTask(tid) {
  if (!confirm('Delete this task and all completion records?')) return;
  var btn     = document.getElementById('db-' + tid);
  if (btn) btn.disabled = true;
  var related = gComps.filter(function(c) { return c.taskId === String(tid); });
  gDel('/sites/' + gSiteId + '/lists/' + gTLId + '/items/' + tid)
  .then(function() {
    return Promise.all(related.map(function(c) {
      return gDel('/sites/' + gSiteId + '/lists/' + gCLId + '/items/' + c.id);
    }));
  }).then(function() {
    gTasks = gTasks.filter(function(t) { return t.id !== tid; });
    gComps = gComps.filter(function(c) { return c.taskId !== String(tid); });
    renderMy(); renderManage();
    showToast('Task deleted.');
  }).catch(function(e) { showToast('Error: ' + e.message, true); if (btn) btn.disabled = false; });
}

function showToast(msg, isErr) {
  var el = document.getElementById('toast'), tx = document.getElementById('toasttxt');
  tx.textContent = msg; el.classList.toggle('err', !!isErr); el.classList.add('show');
  setTimeout(function() { el.classList.remove('show'); }, 4000);
}
