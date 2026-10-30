const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

process.env.TZ = 'America/Santo_Domingo';
const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const plain = value => JSON.parse(JSON.stringify(value));
class Today extends Date {
  constructor(...args) { super(...(args.length ? args : ['2026-10-30T10:00:00-04:00'])); }
}

function backend() {
  const cache = new Map();
  const config = [['key', 'value'], ['GOOGLE_HOLIDAY_EXCLUDED_DATES_JSON', '[]']];
  const manual = [['fecha', 'motivo']];
  const dates = ['2026-11-06', '2026-11-09', '2026-11-06', ...Array.from({ length: 12 }, (_, i) => '2026-12-' + String(i + 1).padStart(2, '0'))];
  let calendarCalls = 0, calendarFails = false, cancelled = 0, writes = 0, revision = 0;
  const configSheet = { getDataRange: () => ({ getValues: () => config }), getRange: row => ({ setValue: value => { config[row - 1][1] = value; } }) };
  const menuSheet = { getDataRange: () => ({ getValues: () => [['header']] }), deleteRow: () => { writes++; } };
  const ctx = vm.createContext({ Date: Today, console: { log() {}, warn() {} } });
  vm.runInContext(source, ctx);
  Object.assign(ctx, {
    getTodayYmd_: () => '2026-10-30',
    formatDate_: date => {
      if (Number.isNaN(date.getTime())) throw new Error('Fecha invalida.');
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    },
    normalizeMenuText_: value => String(value || ''),
    getConfigValue_: key => (config.find(row => row[0] === key) || [null, ''])[1],
    readJsonCache_: key => cache.get(key) || null,
    writeJsonCache_: (key, value) => cache.set(key, value),
    getRevisionValue_: () => revision,
    bumpRevisionValue_: () => { revision++; },
    invalidateMenuDataCache_: () => {}, invalidateUserInitCache_: () => {},
    ensureOperationalConfigKeys_: () => {},
    readSheetValues_: sheet => sheet.getDataRange().getValues(),
    CacheService: { getScriptCache: () => ({ remove: key => cache.delete(key) }) },
    SpreadsheetApp: { getActive: () => ({ getSheetByName: name => name === 'Config' ? configSheet : name === 'Menu' ? menuSheet : { getDataRange: () => ({ getValues: () => manual }) } }) },
    CalendarApp: { getCalendarById: () => {
      if (calendarFails) throw new Error('Permisos insuficientes');
      return { getEvents: () => { calendarCalls++; return dates.map(date => ({ getStartTime: () => new Today(date + 'T00:00:00'), getTitle: () => 'Feriado' })); } };
    } },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    getUserInfo_: () => ({ rol: 'ADMIN_GEN', estado: 'ACTIVO' }),
    getMenuCategoryMap_: () => ({ Arroces: { nombre: 'Arroces', estado: 'ACTIVO' } }),
    cancelActiveOrdersForPlans_: () => { cancelled++; return []; },
    notifyLunchSuspendedCancellations_: () => 0
  });
  return { ctx, config, manual, cache, failCalendar: () => { calendarFails = true; cache.clear(); }, counts: () => ({ calendarCalls, cancelled, writes }) };
}

test('6 de noviembre bloqueado originalmente; excepcion persistente habilita el 6 y conserva el 9', () => {
  const b = backend();
  assert.ok(b.ctx.getHolidaysSet_().has('2026-11-06'));
  const result = b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true);
  assert.equal(result.ok, true);
  assert.equal(result.calendarHolidays.length, 10);
  assert.equal(result.calendarHolidays.filter(h => h.date === '2026-11-06').length, 1);
  assert.equal(result.calendarHolidays[0].excluded, true);
  assert.deepEqual(JSON.parse(b.config[1][1]), ['2026-11-06']);
  assert.equal(b.ctx.getHolidaysSet_().has('2026-11-06'), false);
  assert.ok(b.ctx.getHolidaysSet_().has('2026-11-09'));
  for (let day = 2; day <= 6; day++) assert.ok(b.ctx.isDateAllowedForMenuManagement_('2026-11-0' + day));
  assert.equal(b.ctx.isDateAllowedForMenuManagement_('2026-11-07'), false);
  assert.equal(b.ctx.isDateAllowedForMenuManagement_('2026-10-30'), false);
  assert.equal(b.counts().calendarCalls, 1);
  assert.equal(b.counts().cancelled, 0);
});

test('dia libre manual prevalece; reactivar Google recupera bloqueo y cancela pedidos', () => {
  const b = backend();
  b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true);
  b.manual.push(['2026-11-06', 'Suspension institucional']);
  b.ctx.invalidateHolidayDataCache_();
  assert.ok(b.ctx.getHolidaysSet_().has('2026-11-06'));
  assert.equal(b.ctx.getAdminHolidayData_().calendarHolidays[0].manualBlocked, true);
  b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', false);
  assert.deepEqual(JSON.parse(b.config[1][1]), []);
  assert.equal(b.counts().cancelled, 1);
});

test('autorizacion, fechas, estados y calendario inaccesible no guardan excepciones', () => {
  const b = backend();
  for (const [date, state] of [['2026-02-30', true], ['2026-11-03', true], ['2026-11-06', 'true'], ['2025-11-06', true]]) {
    assert.throws(() => b.ctx.apiSetGoogleHolidayExcluded(date, state));
  }
  b.ctx.getUserInfo_ = () => ({ rol: 'USER', estado: 'ACTIVO' });
  assert.throws(() => b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true), /Denegado/);
  assert.throws(() => b.ctx.apiGetHolidays(), /Denegado/);
  b.ctx.getUserInfo_ = () => ({ rol: 'ADMIN_GEN', estado: 'INACTIVO' });
  assert.throws(() => b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true), /Denegado/);
  b.ctx.getUserInfo_ = () => ({ rol: 'ADMIN_GEN', estado: 'ACTIVO' });
  b.failCalendar();
  assert.throws(() => b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true), /Google Calendar/);
  assert.deepEqual(JSON.parse(b.config[1][1]), []);
});

test('fallo de Google conserva dias manuales, avisa y evita escribir importaciones', () => {
  const b = backend(); b.manual.push(['2026-11-02', 'Libre']); b.failCalendar();
  const state = b.ctx.getAdminHolidayData_();
  assert.equal(state.calendarAvailable, false);
  assert.ok(state.effectiveHolidays.includes('2026-11-02'));
  assert.throws(() => b.ctx.apiSaveWeeklyMenu({ '2026-11-06': [{ cat: 'Arroces', plato: 'Arroz' }] }), /Google Calendar/);
  assert.equal(b.counts().writes, 0);
});

test('importacion bloqueada identifica fecha antes de cualquier escritura', () => {
  const b = backend();
  assert.throws(() => b.ctx.apiSaveWeeklyMenu({ '2026-11-02': [], '2026-11-06': [] }), /2026-11-06.*Calendario/);
  assert.equal(b.counts().writes, 0);
  b.config[1][1] = 'JSON roto';
  assert.ok(b.ctx.getHolidayState_().effectiveHolidays.includes('2026-11-06'));
});

function frontend(b) {
  let ui;
  const runner = {
    withSuccessHandler(fn) { this.success = fn; return this; },
    withFailureHandler(fn) { this.failure = fn; return this; },
    apiGetHolidays() { assert.equal(ui.adminState.loading, true); this.success(b.ctx.apiGetHolidays()); }
  };
  const ctx = vm.createContext({ console, Date: Today, window: {}, google: { script: { run: runner } }, Vue: {
    ref: value => ({ value }), reactive: value => value,
    computed: fn => ({ get value() { return fn(); } }),
    watch() {}, onMounted() {}, nextTick: fn => fn && fn(),
    createApp: app => ({ mount: () => { ui = app.setup(); } })
  } });
  const js = fs.readFileSync(path.join(root, 'js.html'), 'utf8').replace(/^\s*<script[^>]*>/, '').replace(/<\/script>\s*$/, '');
  vm.runInContext(js, ctx);
  ui.menuCategories.value = [{ id: 'Arroces', nombre: 'Arroces', estado: 'ACTIVO', orden: 10 }];
  ui.adminState.menuImportModal.weekStart = '2026-11-02';
  ui.adminState.menuImportModal.text = 'Arroces\nArroz Lunes\tArroz Martes\tArroz Miercoles\tArroz Jueves\tArroz Viernes';
  return ui;
}

test('vista previa usa feriados efectivos y preserva cinco columnas despues de desactivar el 6', () => {
  const b = backend(), ui = frontend(b);
  ui.processMenuImport();
  assert.match(ui.adminState.menuImportModal.error, /2026-11-06.*Calendario/);
  assert.equal(ui.adminState.loading, false);
  b.ctx.apiSetGoogleHolidayExcluded('2026-11-06', true);
  ui.processMenuImport();
  assert.equal(ui.adminState.menuImportModal.error, null);
  assert.equal(ui.adminState.menuImportModal.preview.length, 5);
  assert.deepEqual(plain(ui.adminState.menuImportModal.preview.map(day => day.date)), ['2026-11-02', '2026-11-03', '2026-11-04', '2026-11-05', '2026-11-06']);
  assert.equal(ui.adminState.loading, false);
});
