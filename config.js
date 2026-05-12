// ── Config ───────────────────────────────────────────────
var CFG = {
  clientId:     'f023a1d2-c805-4fd1-80dc-79f4c7c75a50',
  tenantId:     'a3f79504-fd00-417d-9194-0559070d1f8e',
  siteUrl:      'https://bangcreations.sharepoint.com/sites/BANGTEAM',
  tenant:       'bangcreations',
  listTasks:    'HR_Tasks',
  listComps:    'HR_Completions',
  hrGroupId:    'cec65ee4-f9fc-4beb-9b68-2f08f4ec78e6',
  listExpenses: 'expenses_data',
  listMileage:  'mileage_data',
  accGroupId:   'a6ec43fb-a114-4c72-81e5-064fd2080d58',
  hbGroupId:    '2f5f0cf0-9295-4cc0-90e2-b32fe775d011',  // Handbook Editors
  hbFolder:     'Handbook and HR'
};

// Scopes — includes User.ReadBasic.All and GroupMember.Read.All
var SCOPES = [
  'https://graph.microsoft.com/Sites.ReadWrite.All',
  'https://graph.microsoft.com/User.ReadBasic.All',
  'https://graph.microsoft.com/GroupMember.Read.All'
];

// ── State ────────────────────────────────────────────────
var gMsal         = null;
var gUser         = null;   // { name, email, ini }
var gSiteId       = null;
var gTLId         = null;   // HR_Tasks list id
var gCLId         = null;   // HR_Completions list id
var gELId         = null;   // expenses_data list id
var gMLId         = null;   // mileage_data list id
var gTasks        = [];
var gComps        = [];
var gUsers        = [];     // [{ name, email }] — fetched from Graph
var gExpenses     = [];     // standard expenses
var gMileage      = [];     // mileage records
var gIsHrAdmin         = false;
var gIsAccMgr          = false;
var gIsHandbookEditor  = false;
var gHbDriveId         = null;   // Documents drive id
var gHbDocs            = [];     // [{ id, name, tags[], updated, webUrl }]
var gHbTags            = [];     // sorted unique tag strings
var gHbActiveTags      = [];     // currently selected tag filters
var gHbSearch          = '';
var gHbViewing         = null;   // id of open document, or null
var gHbBlobUrl         = null;   // current viewer blob URL (for cleanup)
var gShowAllTasks = false;
var gMyFilter     = 'pending';   // null | 'pending' | 'overdue' | 'complete'
var gExpType      = 'standard'; // active expense log form
var gAccFilter       = 'all';
var gAccFilterState  = 'unprocessed'; // 'all' | 'unprocessed' | 'processed'
var gAccFilterDate   = 'all';         // 'all' | 'this-month' | 'last-month' | 'this-year'
var gAccFilterPerson = null;          // email to filter by, or null for all
var gShowAllExp      = false;
var gEditingExpId    = null;  // id of expense being edited, or null
var gEditingExpType  = null;  // 'standard' | 'mileage' | null
