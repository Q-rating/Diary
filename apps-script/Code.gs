/**
 * 만년 다이어리 ↔ Google 캘린더 · Notion 연동 (Google Apps Script)  — v4
 * 스크립트 속성: DIARY_KEY(필수), NOTION_TOKEN, NOTION_DB(노션 연동 시)
 * 노션 DB 열: 제목(title, 이름 무관), 날짜(date), 카테고리(select), 본문(text = 1~2줄 요약), 시간(text), 장소(text), 슬롯(number, 선택), 첨부(text, 선택)
 * 전체 본문은 각 노션 페이지 안(본문 블록)에 저장됩니다.
 */
var VERSION = '4';
var PROPS = PropertiesService.getScriptProperties();
var NOTION_VERSION = '2022-06-28';
var COL = { date: '날짜', cat: '카테고리', body: '본문', time: '시간', place: '장소', slot: '슬롯', files: '첨부' };

function setup() {
  CalendarApp.getDefaultCalendar().getName();
  if (!PROPS.getProperty('DIARY_KEY')) throw new Error('스크립트 속성에 DIARY_KEY를 추가하세요');
  if (hasNotion_()) { CacheService.getScriptCache().remove('schema'); Logger.log('노션 DB 연결됨: ' + Object.keys(notionSchema_().props).join(', ')); }
  else Logger.log('노션 미설정 — 캘린더만 연동됩니다');
  Logger.log('설정 완료 (코드 v' + VERSION + ')');
}

function doGet() { return json_({ ok: true, app: 'diary', version: VERSION }); }

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (!req.key || req.key !== PROPS.getProperty('DIARY_KEY')) return json_({ ok: false, error: '키가 일치하지 않습니다' });
    var fn = { ping: ping_, pull: pull_, push: push_ }[req.action];
    if (!fn) return json_({ ok: false, error: '알 수 없는 요청' });
    var lock = LockService.getScriptLock(); lock.waitLock(25000);
    try { return json_(fn(req)); } finally { lock.releaseLock(); }
  } catch (err) { return json_({ ok: false, error: String((err && err.message) || err) }); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hasNotion_() { return !!(PROPS.getProperty('NOTION_TOKEN') && PROPS.getProperty('NOTION_DB')); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function fmt_(d, f) { return Utilities.formatDate(d, tz_(), f); }
function day_(s) { return Utilities.parseDate(s, tz_(), 'yyyy-MM-dd'); }

function ping_() {
  var out = { ok: true, version: VERSION, calendar: CalendarApp.getDefaultCalendar().getName(), notion: hasNotion_(), timezone: tz_() };
  if (out.notion) { try { CacheService.getScriptCache().remove('schema'); notionSchema_(); } catch (e) { out.notionError = String(e.message || e); } }
  return out;
}

/* ---------- 가져오기 ---------- */
function pull_(req) {
  var from = day_(req.from), to = new Date(day_(req.to).getTime() + 864e5), events = [];
  CalendarApp.getAllCalendars().forEach(function (c) {
    if (c.isHidden() || !c.isSelected()) return;
    c.getEvents(from, to).forEach(function (ev) {
      var diary = ev.getTag('diary') === '1', title = ev.getTitle(), cal = c.getName();
      if (ev.isAllDayEvent()) {
        for (var d = new Date(ev.getAllDayStartDate()); d < ev.getAllDayEndDate(); d = new Date(d.getTime() + 864e5)) {
          if (d >= from && d < to) events.push({ date: fmt_(d, 'yyyy-MM-dd'), time: '종일', title: title, cal: cal, diary: diary });
        }
      } else {
        events.push({ date: fmt_(ev.getStartTime(), 'yyyy-MM-dd'), time: fmt_(ev.getStartTime(), 'HH:mm'), title: title, cal: cal, diary: diary });
      }
    });
  });
  var out = { ok: true, events: events, notion: null, deleted: [] };
  if (hasNotion_()) {
    try { var r = notionPull_(req.from, req.known || {}); out.notion = r.items; out.deleted = r.deleted; }
    catch (e) { out.notionError = String(e.message || e); }
  }
  return out;
}

/* ---------- 보내기 ---------- */
function push_(req) {
  var results = {}, notion = hasNotion_(), s = null, notionErr = '';
  if (notion) { try { s = notionSchema_(); } catch (e) { notionErr = String(e.message || e); } }
  (req.items || []).forEach(function (it) {
    var r = { notionId: it.notionId || '', eventId: it.eventId || '' }, errs = [];
    if (notion) {
      if (!s) errs.push(notionErr);
      else {
        try { r.notionId = notionUpsert_(s, it); }
        catch (e) {
          if (/property|validation|schema/i.test(String(e))) {
            try { CacheService.getScriptCache().remove('schema'); s = notionSchema_(); r.notionId = notionUpsert_(s, it); }
            catch (e2) { errs.push(String(e2.message || e2)); }
          } else errs.push(String(e.message || e));
        }
      }
    }
    try {
      if (it.cal) r.eventId = eventUpsert_(it);
      else if (it.eventId) { eventDelete_(it.eventId); r.eventId = ''; }
    } catch (e) { errs.push('캘린더: ' + (e.message || e)); }
    if (errs.length) r.error = errs.join(' / ');
    results[it.id] = r;
  });
  if (notion) (req.notionDel || []).forEach(function (id) { try { notion_('patch', '/pages/' + id, { archived: true }); } catch (e) {} });
  (req.calDel || []).forEach(eventDelete_);
  return { ok: true, results: results };
}

/* ---------- Google 캘린더 ---------- */
function eventUpsert_(it) {
  var cal = CalendarApp.getDefaultCalendar(), title = it.title || it.catName, ev = null;
  if (it.eventId) { try { ev = cal.getEventById(it.eventId); } catch (e) {} }
  if (/^\d{2}:\d{2}$/.test(it.time || '')) {
    var st = Utilities.parseDate(it.date + ' ' + it.time, tz_(), 'yyyy-MM-dd HH:mm'), en = new Date(st.getTime() + 36e5);
    if (ev) ev.setTime(st, en); else ev = cal.createEvent(title, st, en);
  } else {
    if (ev) ev.setAllDayDate(day_(it.date)); else ev = cal.createAllDayEvent(title, day_(it.date));
  }
  ev.setTitle(title);
  ev.setDescription('[' + it.catName + '] ' + (it.body || ''));
  ev.setLocation(it.place || '');
  ev.setTag('diary', '1');
  return ev.getId();
}
function eventDelete_(id) { try { var ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) {} }

/* ---------- Notion ---------- */
function notion_(method, path, body) {
  var opt = { method: method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('NOTION_TOKEN'), 'Notion-Version': NOTION_VERSION } };
  if (body) opt.payload = JSON.stringify(body);
  for (var attempt = 0; ; attempt++) {
    var res = UrlFetchApp.fetch('https://api.notion.com/v1' + path, opt), code = res.getResponseCode();
    if ((code === 429 || code >= 500) && attempt < 4) {
      var wait = Number((res.getHeaders() || {})['Retry-After'] || 0) * 1000 || 800 * (attempt + 1);
      Utilities.sleep(Math.min(wait, 5000)); continue;
    }
    var j = {}; try { j = JSON.parse(res.getContentText() || '{}'); } catch (e) {}
    if (code >= 300) throw new Error('Notion: ' + (j.message || code));
    return j;
  }
}
function notionSchema_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('schema');
  if (hit) return JSON.parse(hit);
  var db = notion_('get', '/databases/' + PROPS.getProperty('NOTION_DB')), props = {}, title = null;
  Object.keys(db.properties).forEach(function (n) { props[n] = db.properties[n].type; if (props[n] === 'title') title = n; });
  if (props[COL.date] !== 'date') throw new Error("노션 '" + COL.date + "' 열이 없거나 날짜 유형이 아닙니다");
  if (props[COL.cat] && props[COL.cat] !== 'select' && props[COL.cat] !== 'multi_select') throw new Error("노션 '" + COL.cat + "' 열은 선택 유형이어야 합니다");
  var s = { props: props, title: title }; cache.put('schema', JSON.stringify(s), 300); return s;
}
function text_(p) { return p ? (p.rich_text || p.title || []).map(function (t) { return t.plain_text; }).join('') : ''; }
function summary_(body) {
  var lines = String(body || '').split(/\r?\n/).map(function (l) { return l.replace(/^\s*([-*•]|\d+[.)]|#{1,3})\s*/, '').trim(); }).filter(function (l) { return l; });
  var s = lines.slice(0, 2).join(' / ');
  return s.length > 100 ? s.slice(0, 99) + '…' : s;
}
function toBlocks_(body) {
  var rt = function (t) { return [{ type: 'text', text: { content: t.slice(0, 2000) } }]; };
  return String(body || '').split(/\r?\n/).map(function (l) {
    var m;
    if ((m = l.match(/^\s*[-*•]\s+(.*)$/))) return { object: 'block', type: 'bulleted_list_item', bulleted_list_item: { rich_text: rt(m[1]) } };
    if ((m = l.match(/^\s*\d+[.)]\s+(.*)$/))) return { object: 'block', type: 'numbered_list_item', numbered_list_item: { rich_text: rt(m[1]) } };
    if ((m = l.match(/^#{1,3}\s+(.*)$/))) return { object: 'block', type: 'heading_3', heading_3: { rich_text: rt(m[1]) } };
    return { object: 'block', type: 'paragraph', paragraph: { rich_text: l ? rt(l) : [] } };
  });
}
function readBody_(pageId) {
  var out = [], cursor = null, n = 0;
  do {
    var r = notion_('get', '/blocks/' + pageId + '/children?page_size=100' + (cursor ? '&start_cursor=' + cursor : ''));
    r.results.forEach(function (b) {
      var v = b[b.type] || {}, t = (v.rich_text || []).map(function (x) { return x.plain_text; }).join('');
      if (b.type === 'numbered_list_item') { n++; out.push(n + '. ' + t); return; }
      n = 0;
      if (b.type === 'bulleted_list_item') out.push('- ' + t);
      else if (b.type === 'to_do') out.push((v.checked ? '[x] ' : '[ ] ') + t);
      else if (/^heading_/.test(b.type)) out.push('# ' + t);
      else if (v.rich_text) out.push(t);
    });
    cursor = r.has_more ? r.next_cursor : null;
  } while (cursor);
  return out.join('\n').replace(/\n+$/, '');
}
function writeBody_(pageId, body) {
  for (var i = 0; i < 30; i++) {
    var r = notion_('get', '/blocks/' + pageId + '/children?page_size=100');
    if (!r.results.length) break;
    r.results.forEach(function (b) { notion_('delete', '/blocks/' + b.id); });
    if (!r.has_more) break;
  }
  var blocks = toBlocks_(body);
  for (var j = 0; j < blocks.length; j += 100) notion_('patch', '/blocks/' + pageId + '/children', { children: blocks.slice(j, j + 100) });
}
function pageItem_(s, pg, known) {
  var p = pg.properties, d = p[COL.date] && p[COL.date].date;
  if (!d || !d.start) return null;
  var c = p[COL.cat], cat = c ? (c.select ? c.select.name : (c.multi_select && c.multi_select[0] ? c.multi_select[0].name : '')) : '';
  var time = text_(p[COL.time]); if (!time && d.start.length > 10) time = d.start.slice(11, 16);
  var body = null;
  if (known[pg.id] !== pg.last_edited_time) body = readBody_(pg.id) || text_(p[COL.body]);
  return { id: pg.id, date: d.start.slice(0, 10), title: text_(p[s.title]), cat: cat, body: body, time: time, place: text_(p[COL.place]),
    slot: p[COL.slot] ? p[COL.slot].number : null, edited: pg.last_edited_time };
}
function notionPull_(from, known) {
  var s = notionSchema_(), items = [], seen = {}, deleted = [], cursor = null;
  do {
    var body = { page_size: 100, filter: { property: COL.date, date: { on_or_after: from } } };
    if (cursor) body.start_cursor = cursor;
    var r = notion_('post', '/databases/' + PROPS.getProperty('NOTION_DB') + '/query', body);
    r.results.forEach(function (pg) { seen[pg.id] = 1; var it = pageItem_(s, pg, known); if (it) items.push(it); });
    cursor = r.has_more ? r.next_cursor : null;
  } while (cursor);
  Object.keys(known).forEach(function (id) {
    if (seen[id]) return;
    try {
      var pg = notion_('get', '/pages/' + id);
      if (pg.archived || pg.in_trash) deleted.push(id);
      else { var it = pageItem_(s, pg, known); if (it) items.push(it); }
    } catch (e) { if (/Could not find|object_not_found/i.test(String(e))) deleted.push(id); }
  });
  return { items: items, deleted: deleted };
}
function notionProps_(s, it) {
  var P = {}, has = function (n, t) { return s.props[n] && (!t || s.props[n] === t); },
      rt = function (v) { return [{ text: { content: String(v || '').slice(0, 2000) } }]; };
  P[s.title] = { title: rt(it.title || it.catName) };
  P[COL.date] = { date: { start: it.date } };
  if (has(COL.cat, 'select')) P[COL.cat] = { select: { name: it.catName } };
  else if (has(COL.cat, 'multi_select')) P[COL.cat] = { multi_select: [{ name: it.catName }] };
  if (has(COL.body, 'rich_text')) P[COL.body] = { rich_text: rt(summary_(it.body)) };
  if (has(COL.time, 'rich_text')) P[COL.time] = { rich_text: rt(it.time) };
  if (has(COL.place, 'rich_text')) P[COL.place] = { rich_text: rt(it.place) };
  if (has(COL.slot, 'number')) P[COL.slot] = { number: it.slot };
  if (has(COL.files, 'rich_text')) P[COL.files] = { rich_text: rt((it.files || []).join(', ')) };
  return P;
}
function findTwin_(s, it) {
  var title = it.title || it.catName;
  var r = notion_('post', '/databases/' + PROPS.getProperty('NOTION_DB') + '/query', { page_size: 5,
    filter: { and: [{ property: COL.date, date: { equals: it.date } }, { property: s.title, title: { equals: title } }] } });
  return r.results.length ? r.results[0].id : '';
}
function notionUpsert_(s, it) {
  var properties = notionProps_(s, it), id = it.notionId || findTwin_(s, it);
  if (id) {
    var ok = false;
    try { notion_('patch', '/pages/' + id, { properties: properties, archived: false }); ok = true; } catch (e) { if (!/Could not find|object_not_found|archived/i.test(String(e))) throw e; }
    if (ok) { writeBody_(id, it.body); return id; }
  }
  var blocks = toBlocks_(it.body);
  id = notion_('post', '/pages', { parent: { database_id: PROPS.getProperty('NOTION_DB') }, properties: properties, children: blocks.slice(0, 100) }).id;
  for (var j = 100; j < blocks.length; j += 100) notion_('patch', '/blocks/' + id + '/children', { children: blocks.slice(j, j + 100) });
  return id;
}
