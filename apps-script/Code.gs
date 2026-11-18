/**
 * 만년 다이어리 ↔ Google 캘린더 · Notion 연동 (Google Apps Script)
 * 스크립트 속성: DIARY_KEY(필수), NOTION_TOKEN, NOTION_DB(노션 연동 시)
 * 노션 DB 열: 제목(title, 이름 무관), 날짜(date), 카테고리(select), 본문(text), 시간(text), 슬롯(number, 선택), 첨부(text, 선택)
 */
var PROPS = PropertiesService.getScriptProperties();
var NOTION_VERSION = '2022-06-28';
var COL = { date: '날짜', cat: '카테고리', body: '본문', time: '시간', slot: '슬롯', files: '첨부' };

function setup() {
  CalendarApp.getDefaultCalendar().getName();
  if (!PROPS.getProperty('DIARY_KEY')) throw new Error('스크립트 속성에 DIARY_KEY를 추가하세요');
  if (hasNotion_()) Logger.log('노션 DB 연결됨: ' + Object.keys(notionSchema_().props).join(', '));
  else Logger.log('노션 미설정 — 캘린더만 연동됩니다');
  Logger.log('설정 완료');
}

function doGet() { return json_({ ok: true, app: 'diary' }); }

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (!req.key || req.key !== PROPS.getProperty('DIARY_KEY')) return json_({ ok: false, error: '키가 일치하지 않습니다' });
    var fn = { ping: ping_, pull: pull_, push: push_ }[req.action];
    if (!fn) return json_({ ok: false, error: '알 수 없는 요청' });
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return json_(fn(req)); } finally { lock.releaseLock(); }
  } catch (err) { return json_({ ok: false, error: String((err && err.message) || err) }); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hasNotion_() { return !!(PROPS.getProperty('NOTION_TOKEN') && PROPS.getProperty('NOTION_DB')); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function fmt_(d, f) { return Utilities.formatDate(d, tz_(), f); }
function day_(s) { return Utilities.parseDate(s, tz_(), 'yyyy-MM-dd'); }

function ping_() { return { ok: true, calendar: CalendarApp.getDefaultCalendar().getName(), notion: hasNotion_() }; }

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
  return { ok: true, events: events, notion: hasNotion_() ? notionPull_(req.from) : null };
}

function push_(req) {
  var results = {}, notion = hasNotion_(), s = notion ? notionSchema_() : null;
  (req.items || []).forEach(function (it) {
    var r = { notionId: it.notionId || '', eventId: it.eventId || '' };
    if (notion) r.notionId = notionUpsert_(s, it);
    if (it.cal) r.eventId = eventUpsert_(it);
    else if (it.eventId) { eventDelete_(it.eventId); r.eventId = ''; }
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
  ev.setTag('diary', '1');
  return ev.getId();
}
function eventDelete_(id) { try { var ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) {} }

/* ---------- Notion ---------- */
function notion_(method, path, body) {
  var opt = { method: method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('NOTION_TOKEN'), 'Notion-Version': NOTION_VERSION } };
  if (body) opt.payload = JSON.stringify(body);
  var res = UrlFetchApp.fetch('https://api.notion.com/v1' + path, opt), j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) throw new Error('Notion: ' + (j.message || res.getResponseCode()));
  return j;
}
function notionSchema_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('schema');
  if (hit) return JSON.parse(hit);
  var db = notion_('get', '/databases/' + PROPS.getProperty('NOTION_DB')), props = {}, title = null;
  Object.keys(db.properties).forEach(function (n) { props[n] = db.properties[n].type; if (props[n] === 'title') title = n; });
  var s = { props: props, title: title }; cache.put('schema', JSON.stringify(s), 600); return s;
}
function text_(p) { return p ? (p.rich_text || p.title || []).map(function (t) { return t.plain_text; }).join('') : ''; }
function notionPull_(from) {
  var s = notionSchema_(), out = [], cursor = null;
  do {
    var body = { page_size: 100, filter: { property: COL.date, date: { on_or_after: from } } };
    if (cursor) body.start_cursor = cursor;
    var r = notion_('post', '/databases/' + PROPS.getProperty('NOTION_DB') + '/query', body);
    r.results.forEach(function (pg) {
      var p = pg.properties, d = p[COL.date] && p[COL.date].date;
      if (!d || !d.start) return;
      var c = p[COL.cat], cat = c ? (c.select ? c.select.name : (c.multi_select && c.multi_select[0] ? c.multi_select[0].name : '')) : '';
      var time = text_(p[COL.time]); if (!time && d.start.length > 10) time = d.start.slice(11, 16);
      out.push({ id: pg.id, date: d.start.slice(0, 10), title: text_(p[s.title]), cat: cat, body: text_(p[COL.body]), time: time,
        slot: p[COL.slot] ? p[COL.slot].number : null, edited: pg.last_edited_time });
    });
    cursor = r.has_more ? r.next_cursor : null;
  } while (cursor);
  return out;
}
function notionProps_(s, it) {
  var P = {}, has = function (n, t) { return s.props[n] && (!t || s.props[n] === t); },
      rt = function (v) { return [{ text: { content: String(v || '').slice(0, 2000) } }]; };
  P[s.title] = { title: rt(it.title || it.catName) };
  if (has(COL.date, 'date')) P[COL.date] = { date: { start: it.date } };
  if (has(COL.cat, 'select')) P[COL.cat] = { select: { name: it.catName } };
  else if (has(COL.cat, 'multi_select')) P[COL.cat] = { multi_select: [{ name: it.catName }] };
  if (has(COL.body, 'rich_text')) P[COL.body] = { rich_text: rt(it.body) };
  if (has(COL.time, 'rich_text')) P[COL.time] = { rich_text: rt(it.time) };
  if (has(COL.slot, 'number')) P[COL.slot] = { number: it.slot };
  if (has(COL.files, 'rich_text')) P[COL.files] = { rich_text: rt((it.files || []).join(', ')) };
  return P;
}
function notionUpsert_(s, it) {
  var properties = notionProps_(s, it);
  if (it.notionId) { try { notion_('patch', '/pages/' + it.notionId, { properties: properties }); return it.notionId; } catch (e) {} }
  return notion_('post', '/pages', { parent: { database_id: PROPS.getProperty('NOTION_DB') }, properties: properties }).id;
}
