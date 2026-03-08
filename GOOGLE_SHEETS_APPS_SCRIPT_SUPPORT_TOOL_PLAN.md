# OVO Utilities Support Tool — Google Sheets + Apps Script Only (No GitHub, No Hosting, No 3rd Party)

## Aapki condition ke hisaab se final answer
Haan, ye **100% possible** hai sirf:
- Google Sheet
- Bound Google Apps Script
- Google Workspace login

Aur **nahi chahiye**:
- GitHub deploy
- External hosting (VPS, Firebase Hosting, Render, etc.)
- Koi 3rd-party app/tool

---

## Final architecture (strictly internal)

1. **Single Google Spreadsheet** (database ki tarah)
2. **Apps Script (bound to same sheet)**
3. **Apps Script Web App deployment**
   - Execute as: `User accessing the web app`
   - Access: `Only users in your domain` (ya specific users)
4. **Role-based UI** same web app mein
   - Agent dashboard
   - SME dashboard
   - Manager dashboard

> Is model mein deployment Google ke andar hi hota hai. Koi external hosting nahi lagti.

---

## Sheet structure (tabs)

### 1) `Users`
| email | name | role | is_active |
|---|---|---|---|
| a@ovo.com | Agent A | agent | TRUE |
| s@ovo.com | SME S | sme | TRUE |
| m@ovo.com | Manager M | manager | TRUE |

### 2) `Queries`
| query_id | created_at | agent_email | customer_ref | issue_type | priority | status | assigned_sme | updated_at | resolved_at |

Status values:
- `open`
- `assigned`
- `in_progress`
- `resolved`
- `closed`

### 3) `SME_Status`
| sme_email | status | reason | last_updated | active_load |

Status values:
- `available`
- `busy`
- `break`
- `offline`

### 4) `Messages`
| msg_id | query_id | sender_email | sender_role | message | sent_at |

### 5) `Audit_Log`
| event_id | event_type | actor_email | query_id | details | created_at |

### 6) `Manager_Metrics`
Precomputed metrics for fast manager dashboard.

---

## Dashboard requirements mapping

## 1) Agent Dashboard (Care Team)
- Raise Query form
- SME counts: available / busy / break
- My Queries (sirf logged-in agent ki)
- Query-wise chat (WhatsApp style)

## 2) SME Dashboard (Billing Team)
- My Assigned Queries
- Status buttons: Available / Busy / Break
- Agent ke saath query chat
- Actions: Start, Resolve, Ask More Info

## 3) Manager Dashboard
- Sab SME status live
- Kis SME ne kitne issues handle kiye
- Queue health: open, assigned, in progress, overdue
- Break pe kaun hai / available kaun hai

---

## Extremely fast banane ka practical plan

1. **Full sheet scans avoid karo**
   - `getDataRange()` har request pe mat chalao.
2. **Indexed reads karo**
   - Header map + exact column range use karo.
3. **Chat append-only rakho**
   - Messages tab mein sirf append row.
4. **Incremental fetch**
   - `lastSeenTimestamp` ke baad ki messages only.
5. **CacheService use karo**
   - Role lookup, SME counts, recent query list.
6. **LockService use karo**
   - Assignment / status transitions race condition se bachao.
7. **Manager metrics precompute**
   - Time trigger every 1 min.

Target speed:
- Dashboard first load: `< 1.5s`
- Raise query: `< 800ms`
- Send chat: `< 500ms`
- New chat visibility: `2–3s` polling

---

## Apps Script project structure (same bound project)

- `Code.gs` (router + doGet)
- `Auth.gs` (user/role checks)
- `Queries.gs` (create/assign/update)
- `Chat.gs` (send/fetch)
- `Status.gs` (SME state)
- `Metrics.gs` (manager precompute)
- `Index.html` (shell)
- `Agent.html`
- `SME.html`
- `Manager.html`
- `Styles.html` (modern green UI)
- `ClientJS.html`

---

## Minimal working backend (starter)

### `Code.gs`
```javascript
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OVO Support Tool');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSessionUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('User email unavailable. Domain policy check karein.');
  return { email };
}
```

### `Auth.gs`
```javascript
function getMyProfile() {
  const { email } = getSessionUser();
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const iEmail = head.indexOf('email');
  const iName = head.indexOf('name');
  const iRole = head.indexOf('role');
  const iActive = head.indexOf('is_active');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][iEmail]).toLowerCase() === email.toLowerCase() && data[r][iActive] === true) {
      return { email, name: data[r][iName], role: data[r][iRole] };
    }
  }
  throw new Error('User not authorized');
}
```

### `Queries.gs`
```javascript
function raiseQuery(payload) {
  const me = getMyProfile();
  if (me.role !== 'agent') throw new Error('Only agent can raise query');

  const lock = LockService.getScriptLock();
  lock.waitLock(3000);
  try {
    const ss = SpreadsheetApp.getActive();
    const qSh = ss.getSheetByName('Queries');
    const now = new Date();
    const id = 'Q-' + Utilities.formatString('%06d', qSh.getLastRow());

    const assignedSme = pickAvailableSme_();
    const status = assignedSme ? 'assigned' : 'open';

    qSh.appendRow([
      id, now, me.email,
      payload.customer_ref || '',
      payload.issue_type || '',
      payload.priority || 'P3',
      status,
      assignedSme || '',
      now,
      ''
    ]);

    if (assignedSme) incrementSmeLoad_(assignedSme, 1);
    return { ok: true, query_id: id, assigned_sme: assignedSme || null };
  } finally {
    lock.releaseLock();
  }
}
```

### `Chat.gs`
```javascript
function sendMessage(queryId, text) {
  const me = getMyProfile();
  if (!text || !text.trim()) throw new Error('Empty message');

  const sh = SpreadsheetApp.getActive().getSheetByName('Messages');
  const now = new Date();
  const msgId = 'M-' + Utilities.getUuid();

  sh.appendRow([msgId, queryId, me.email, me.role, text.trim(), now]);
  touchQueryUpdatedAt_(queryId, now);
  return { ok: true, msg_id: msgId, sent_at: now };
}

function getMessages(queryId, afterTs) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Messages');
  const rows = sh.getDataRange().getValues();
  const out = [];
  const after = afterTs ? new Date(afterTs).getTime() : 0;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] !== queryId) continue;
    const ts = new Date(rows[i][5]).getTime();
    if (ts > after) {
      out.push({
        msg_id: rows[i][0],
        sender_email: rows[i][2],
        sender_role: rows[i][3],
        message: rows[i][4],
        sent_at: rows[i][5]
      });
    }
  }
  return out;
}
```

> Upar starter code hai; production mein helper functions (`pickAvailableSme_`, `incrementSmeLoad_`, etc.) add karne honge.

---

## UI design guidance (modern green)

Palette:
- Primary dark: `#14532D`
- Primary: `#16A34A`
- Accent: `#22C55E`
- Background tint: `#DCFCE7`

Rules:
- Card-based clean layout
- Rounded corners + subtle shadow
- Compact table + sticky top stats
- Chat bubbles with time label
- Mobile-friendly responsive grid

---

## Step-by-step build plan (without GitHub)

## Day 1 — Setup + security
1. Sheet create karo + all tabs.
2. Users tab mein real users add karo.
3. Apps Script open (Extensions > Apps Script).
4. `doGet`, `include`, `getMyProfile` setup.
5. Web app deploy (internal users only).

## Day 2 — Agent flow
1. Raise Query form.
2. Query ID generation.
3. Auto assignment to available SME.
4. Agent “My Queries” list.

## Day 3 — SME flow
1. “My Assigned Queries” view.
2. Status toggle (available/break/busy).
3. Resolve/start actions.

## Day 4 — Chat + manager
1. Query chat send/fetch.
2. Manager live overview cards.
3. Daily metrics tab auto update.

## Day 5 — Hardening
1. Role-based access lock down.
2. Audit log everywhere.
3. Performance tuning + UAT.
4. Go-live.

---

## Important deployment settings (must)

Apps Script → Deploy → New Deployment → Web App:
- Execute as: **User accessing the web app**
- Who has access: **Only within your organization**

Isse har action real logged-in user ke context mein chalega, jo aapke use-case ke liye best hai.

---

## Common mistakes avoid karo

- Single giant sheet use mat karo (tab-wise split rakho).
- Har request pe full scan mat karo.
- Chat edits allow mat karo (append-only).
- Role checks skip mat karo.
- Manager dashboard ke liye live heavy compute mat karo (precompute).

---

## Final confirmation
Aapka required system (3 dashboards + query raise + assignment + status + chat + manager visibility) **Google Sheets + Apps Script only** se ban jayega, bina GitHub hosting aur bina third-party tools ke.
