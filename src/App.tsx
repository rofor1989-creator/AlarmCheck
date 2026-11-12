import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { createPortal } from 'react-dom';
import {
  OutageEvent, Oscillogram, EMPTY_EVENT, COLUMNS, VOLTAGE_OPTIONS, REQUIRED_FIELDS,
  isIncomplete, parseDateStr, nowDateStr, getYear,
} from './types';
import { parseWorkbook, exportToExcel, ImportColumn, ParsedSheet } from './excelUtils';
import {
  Plus, Upload, Download, Save, LogOut, Trash2, Filter, X, FileText, AlertTriangle,
  FileJson, FolderOpen, BarChart3, Table as TableIcon,
} from 'lucide-react';

// Версия 3 — пересборка для валидации ветки (формат дат дд.мм.гг чч:мм)
const STORAGE_KEY = 'outage_program_events_v14';

// ---------- Глобальный тип Chart.js (подключается через CDN в index.html) ----------
declare global {
  interface Window { Chart?: any }
}

// ---------- Анализ текста "Анализ работы" ("П-2", "НП-N", "У/НУ АПВ", отказ) ----------
interface AnalysisCounts {
  correct: number;    // "П" / "П-N" — правильно
  incorrect: number;  // "НП" / "НП-N" — неправильно
  arSuccess: number;  // "У" без "НУ" — успешность АРВ/АПВ
  arFail: number;     // "НУ" — неуспешность АРВ/АПВ
  refusal: number;    // метка "Отказ" события
}

const EMPTY_ANALYSIS: AnalysisCounts = { correct: 0, incorrect: 0, arSuccess: 0, arFail: 0, refusal: 0 };

function parseAnalysis(text: string, isRefusal: boolean): AnalysisCounts {
  const res: AnalysisCounts = { ...EMPTY_ANALYSIS };
  const t = String(text ?? '').toUpperCase().replace(/Ё/g, 'Е');
  if (/НП\s*-?\s*\d*/.test(t)) res.incorrect++;        // НП / НП-N — неправильно
  else if (/П\s*-?\s*\d*/.test(t)) res.correct++;      // П / П-N — правильно
  if (/НУ/.test(t)) res.arFail++;                      // НУ — неуспешно
  else if (/\bU\b|У/.test(t.replace(/НУ/g, ''))) res.arSuccess++; // "У" без "НУ" — успешно
  if (isRefusal) res.refusal++;                         // событие помечено как "Отказ"
  return res;
}

// Подвкладки графиков на вкладке "Анализ"
type ChartTab = 'outages' | 'protection' | 'arv' | 'refusal' | 'summary';

const CHART_TABS: { id: ChartTab; label: string }[] = [
  { id: 'outages', label: 'Отключения' },
  { id: 'protection', label: 'Работа защит' },
  { id: 'arv', label: 'Работа АПВ' },
  { id: 'refusal', label: 'Отказы' },
  { id: 'summary', label: 'Сводная' },
];

// ---------- Вкладка "Анализ": 5 подвкладок с графиками Chart.js ----------
function AnalysisTab({ events }: { events: OutageEvent[] }) {
  const [mode, setMode] = useState<'year' | 'month'>('year');
  const [activeChart, setActiveChart] = useState<ChartTab>('outages');
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const chartRef = useRef<any>(null);

  // группировка событий по периодам + подсчёт категорий анализа
  const periods = useMemo(() => {
    const map = new Map<string, AnalysisCounts & { total: number }>();
    const keyOf = (e: OutageEvent): string | null => {
      const d = parseDateStr(e.date);
      if (!d) return null;
      if (mode === 'year') return String(d.getFullYear());
      // формат метки "ММ.ГГГГ" — для хронологической сортировки слева направо
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      return `${mm}.${d.getFullYear()}`;
    };
    for (const e of events) {
      const k = keyOf(e);
      if (!k) continue;
      const cur = map.get(k) ?? { ...EMPTY_ANALYSIS, total: 0 };
      const c = parseAnalysis(e.analysis, e.refusal);
      cur.total++;
      cur.correct += c.correct;
      cur.incorrect += c.incorrect;
      cur.arSuccess += c.arSuccess;
      cur.arFail += c.arFail;
      cur.refusal += c.refusal;
      map.set(k, cur);
    }
    // 1) группировка по периоду
    const keys = [...map.keys()];

    // 2) явная хронологическая сортировка меток (слева направо)
    const cmp = (a: string, b: string): number => {
      if (mode === 'year') return Number(a) - Number(b);            // По годам: возрастание
      const [mA, yA] = a.split('.').map(Number);                    // По месяцам: "ММ.ГГГГ"
      const [mB, yB] = b.split('.').map(Number);
      if (yA !== yB) return yA - yB;
      return mA - mB;
    };
    const sortedIndices = keys
      .map((label, index) => ({ label, index }))
      .sort((x, y) => cmp(x.label, y.label))
      .map((item) => item.index);

    // 3) синхронно пересобираем периоды в отсортированном порядке
    const orderedKeys = sortedIndices.map((i) => keys[i]);
    return orderedKeys.map((k) => ({ label: k, data: map.get(k)! }));
  }, [events, mode]);

  useEffect(() => {
    const ChartLib = window.Chart;
    if (!ChartLib || !canvasRef.current) return;
    if (chartRef.current) { chartRef.current.destroy(); chartRef.current = null; }

    const rawLabels = periods.map((p) => p.label);
    const lineDs = (label: string, data: number[], color: string) => ({
      label,
      data,
      borderColor: color,
      backgroundColor: color,
      pointBackgroundColor: color,
      pointRadius: 4,
      borderWidth: 2,
      tension: 0.25,
      fill: false,
    });
    const barDs = (label: string, data: number[], color: string, stack?: string) => ({
      label,
      data,
      backgroundColor: color,
      ...(stack ? { stack } : {}),
      borderRadius: 3,
      barPercentage: 0.8,
    });

    let type: 'line' | 'bar' = 'bar';
    let datasets: any[] = [];

    switch (activeChart) {
      case 'outages': // График 1: линия количества отключений (синий)
        type = 'line';
        datasets = [lineDs('Отключения', periods.map((p) => p.data.total), '#3b82f6')];
        break;
      case 'protection': // График 2: составной столбик П (зелёный) + НП (красный)
        type = 'bar';
        datasets = [
          barDs('П — правильно', periods.map((p) => p.data.correct), '#10b981', 'prot'),
          barDs('НП — неправильно', periods.map((p) => p.data.incorrect), '#ef4444', 'prot'),
        ];
        break;
      case 'arv': // График 3: составной столбик У (зелёный) + НУ (красный)
        type = 'bar';
        datasets = [
          barDs('У — успешно', periods.map((p) => p.data.arSuccess), '#10b981', 'arv'),
          barDs('НУ — неуспешно', periods.map((p) => p.data.arFail), '#ef4444', 'arv'),
        ];
        break;
      case 'refusal': // График 4: линия отказов (оранжевый)
        type = 'line';
        datasets = [lineDs('Отказы', periods.map((p) => p.data.refusal), '#f59e0b')];
        break;
      case 'summary': // График 5 "Сводная": 6 серий (столбики), визуально 4 параметра рядом
        type = 'bar';
        datasets = [
          barDs('Отключения', periods.map((p) => p.data.total), '#3b82f6', 'outages'),
          barDs('П', periods.map((p) => p.data.correct), '#10b981', 'protection'),
          barDs('НП', periods.map((p) => p.data.incorrect), '#ef4444', 'protection'),
          barDs('У', periods.map((p) => p.data.arSuccess), '#10b981', 'arv'),
          barDs('НУ', periods.map((p) => p.data.arFail), '#ef4444', 'arv'),
          barDs('Отказы', periods.map((p) => p.data.refusal), '#f59e0b', 'refusal'),
        ].map((ds) => ({ ...ds, barPercentage: 0.6, categoryPercentage: 0.8 }));
        break;
    }

    // Явная хронологическая сортировка меток и синхронная пересборка данных серий
    // (единый универсальный механизм для всех пяти подвкладок)
    const cmpLabel = (a: string, b: string): number => {
      if (mode === 'year') return Number(a) - Number(b);            // По годам: возрастание
      const [mA, yA] = a.split('.').map(Number);                    // По месяцам: "ММ.ГГГГ"
      const [mB, yB] = b.split('.').map(Number);
      if (yA !== yB) return yA - yB;
      return mA - mB;
    };
    const sortedIndices = rawLabels
      .map((label, index) => ({ label, index }))
      .sort((x, y) => cmpLabel(x.label, y.label))
      .map((item) => item.index);

    const labels = sortedIndices.map((i) => rawLabels[i]);
    const periodsSorted = sortedIndices.map((i) => periods[i]);     // для tooltip в правильном порядке
    const sortedDatasets = datasets.map((dataset) => ({
      ...dataset,
      data: sortedIndices.map((i) => dataset.data[i]),
    }));

    // Общий tooltip для всех пяти графиков: полная статистика за период
    const tooltipFooter = (items: any[]) => {
      if (!items.length) return '';
      const d = periodsSorted[items[0].dataIndex]?.data;
      if (!d) return '';
      return [
        '————————————',
        `Всего отключений: ${d.total}`,
        `Правильная работа защит: ${d.correct}`,
        `Неправильная работа: ${d.incorrect}`,
        `Отказы: ${d.refusal}`,
        `Успешное АПВ: ${d.arSuccess}`,
        `Неуспешное АПВ: ${d.arFail}`,
      ];
    };

    const isBarStacked = type === 'bar';

    chartRef.current = new ChartLib(canvasRef.current, {
      type,
      data: { labels, datasets: sortedDatasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: activeChart === 'summary' ? 'index' : 'nearest', intersect: activeChart !== 'summary' },
        plugins: {
          legend: { position: 'top' },
          tooltip: { callbacks: { footer: tooltipFooter } },
        },
        scales: {
          x: { ...(isBarStacked ? { stacked: true } : {}) },
          y: {
            beginAtZero: true,
            ...(isBarStacked ? { stacked: true } : {}),
            ticks: { stepSize: 1, precision: 0 },
          },
        },
      },
    });
    return () => {
      if (chartRef.current) { chartRef.current.destroy(); chartRef.current = null; }
    };
  }, [periods, activeChart, mode]);

  return (
    <div className="bg-white rounded-lg shadow border p-4">
      <div className="flex flex-wrap items-center gap-3 mb-3">
        <h2 className="text-base font-bold flex items-center gap-2"><BarChart3 size={18} /> Анализ работы РЗА / АПВ</h2>
        <label className="text-sm font-medium text-slate-600 flex items-center gap-2">
          Период:
          <select
            value={mode}
            onChange={(e) => setMode(e.target.value as 'year' | 'month')}
            className="border border-slate-300 rounded px-2 py-1.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            <option value="year">По годам</option>
            <option value="month">По месяцам</option>
          </select>
        </label>
        <span className="text-xs text-slate-500 ml-auto">Наведите на столбец/точку — в тултипе отобразится полная статистика за период.</span>
      </div>

      {/* Подвкладки графиков */}
      <div className="flex flex-wrap items-center gap-1 mb-4 border-b border-slate-200">
        {CHART_TABS.map((t) => (
          <button key={t.id} onClick={() => setActiveChart(t.id)}
            className={`px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
              activeChart === t.id ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
            }`}>
            {t.label}
          </button>
        ))}
      </div>

      {periods.length === 0 ? (
        <div className="py-16 text-center text-slate-400 text-sm">Нет событий с корректной датой для построения графика.</div>
      ) : (
        <div style={{ height: 420 }}>
          <canvas ref={canvasRef} />
        </div>
      )}
    </div>
  );
}

function loadEvents(): OutageEvent[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch { /* ignore */ }
  return [];
}

type QuickFilter = 'none' | 'refusal' | 'v330' | 'v110' | 'incomplete';

const QUICK_FILTERS: { id: QuickFilter; label: string }[] = [
  { id: 'refusal', label: 'Отключения с отказом' },
  { id: 'v330', label: 'Отключения по сети 330 кВ' },
  { id: 'v110', label: 'Отключения по сети 110 кВ' },
  { id: 'incomplete', label: 'Не оформленные отключения' },
];

// ---------- Модальное окно подтверждения ----------
function ConfirmModal({
  title, buttons, onClose,
}: { title: string; buttons: { label: string; onClick: () => void; danger?: boolean }[]; onClose: () => void }) {
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={onClose}>
      <div className="bg-white rounded-lg shadow-xl p-6 max-w-md w-full mx-4" onClick={(e) => e.stopPropagation()}>
        <div className="text-lg font-semibold mb-4">{title}</div>
        <div className="flex flex-wrap gap-2 justify-end">
          {buttons.map((b) => (
            <button key={b.label} onClick={b.onClick}
              className={`px-4 py-2 rounded font-medium transition-colors ${
                b.danger ? 'bg-red-600 text-white hover:bg-red-700' : 'bg-slate-200 hover:bg-slate-300'
              }`}>
              {b.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ---------- Форма события ----------
function EventForm({
  initial, onSave, onClose,
}: { initial: OutageEvent | null; onSave: (e: OutageEvent) => void; onClose: () => void }) {
  const [ev, setEv] = useState<OutageEvent>(initial ?? { ...EMPTY_EVENT, id: crypto.randomUUID(), oscillograms: [] });
  const oscInputRef = useRef<HTMLInputElement>(null);

  const set = (k: keyof OutageEvent, v: any) => setEv((p) => ({ ...p, [k]: v }));

  // нормализация номера при вводе (только цифры)
  const numStr = ev.num == null ? '' : String(ev.num);

  // Осциллограммы хранятся как относительные ссылки: файл кладётся в папку "Осциллограммы"
  // рядом с html-страницей; в ячейке указывается только имя файла.
  const oscUrl = (o: Oscillogram) => encodeURI(`Осциллограммы/${o.fileName}`);

  const addOscFiles = (files: FileList | null) => {
    if (!files) return;
    Array.from(files).forEach((f) => {
      setEv((p) => ({
        ...p,
        oscillograms: [...p.oscillograms, { id: crypto.randomUUID(), name: f.name, fileName: f.name, type: f.type }],
      }));
    });
  };

  const field = (label: string, key: keyof OutageEvent, placeholder = '') => {
    const multiline = ['notes', 'cause', 'analysis', 'rzaAdjacent', 'rzaPs1', 'rzaPs2', 'ompCalc', 'ompReal'].includes(String(key));
    return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-600">{label}</span>
      {multiline ? (
        <textarea
          rows={3}
          className="mt-1 w-full border border-slate-300 rounded px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          value={(ev as any)[key] ?? ''}
          onChange={(e) => set(key, e.target.value)}
          placeholder={placeholder}
        />
      ) : (
      <input
        className="mt-1 w-full border border-slate-300 rounded px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        value={(ev as any)[key] ?? ''}
        onChange={(e) => set(key, e.target.value)}
        placeholder={placeholder}
      />
      )}
    </label>
    );
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-start justify-center z-50 overflow-y-auto py-8" onClick={onClose}>
      <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-3xl mx-4" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-bold">{initial ? 'Редактирование события' : 'Новое событие'}</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={20} /></button>
        </div>
        <p className="text-xs text-slate-500 mb-4">Все поля необязательны — информацию можно дополнить позже.</p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="block">
            <span className="text-xs font-semibold text-slate-600">Номер события</span>
            <input
              className="mt-1 w-full border border-slate-300 rounded px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              value={numStr}
              onChange={(e) => {
                const digits = e.target.value.replace(/[^0-9]/g, '');
                set('num', digits === '' ? undefined : parseInt(digits, 10));
              }}
              placeholder="авто"
            />
          </label>
          {field('Дата отключения (дд.мм.гг чч:мм)', 'date', nowDateStr())}
          {field('Линия ("ПС №1" - "ПС №2")', 'line', 'ПС 110 кВ "Заря" - ПС 110 кВ "Искра"')}
          <label className="block">
            <span className="text-xs font-semibold text-slate-600">Класс напряжения, кВ</span>
            <select
              className="mt-1 w-full border border-slate-300 rounded px-2 py-1.5 text-sm bg-white"
              value={ev.voltage}
              onChange={(e) => set('voltage', e.target.value)}
            >
              <option value="">— не указан —</option>
              {VOLTAGE_OPTIONS.map((v) => <option key={v} value={v}>{v} кВ</option>)}
              <option value="other">Другое</option>
            </select>
          </label>
          {(ev.voltage === 'other' || (ev.voltage && !VOLTAGE_OPTIONS.includes(ev.voltage))) &&
            field('Напряжение, кВ (уточнение)', 'voltage')}
          {field('Работа РЗА на ПС №1', 'rzaPs1', 'ДЗЛ, АПВ У')}
          {field('Работа РЗА на ПС №2', 'rzaPs2', 'ДЗЛ, АПВ НУ')}
          {field('Причина повреждения', 'cause', 'Обрыв провода в пролёте')}
          {field('ОМП расчетное (РАС/ФИП/МП терминалы + модель)', 'ompCalc', 'РАС Парма, 12,4 км от ПС №1')}
          {field('ОМП реальное (привязка к ПС)', 'ompReal', '12,1 км от ПС №1')}
          {field('Анализ работы (П-2, У/НУ)', 'analysis', 'П-2, АПВ У')}
          {field('Работа РЗА на смежных ПС', 'rzaAdjacent', 'Без замечаний')}
          <label className="block sm:col-span-2">
            <span className="text-xs font-semibold text-slate-600">Примечания</span>
            <textarea
              className="mt-1 w-full border border-slate-300 rounded px-2 py-1.5 text-sm min-h-[60px]"
              value={ev.notes}
              onChange={(e) => set('notes', e.target.value)}
            />
          </label>
        </div>

        <label className="flex items-center gap-2 mt-3 select-none cursor-pointer">
          <input type="checkbox" checked={ev.refusal} onChange={(e) => set('refusal', e.target.checked)}
            className="w-4 h-4 accent-red-600" />
          <span className="text-sm font-semibold text-red-700">Отказ</span>
          <span className="text-xs text-slate-500">(событие будет помечено красным)</span>
        </label>

        <div className="mt-4 border-t pt-3">
          <div className="text-xs font-semibold text-slate-600 mb-2">Осциллограммы</div>
          <input ref={oscInputRef} type="file" multiple accept="image/*,.pdf" className="hidden"
            onChange={(e) => { addOscFiles(e.target.files); e.target.value = ''; }} />
          <button type="button" onClick={() => oscInputRef.current?.click()}
            className="text-sm px-3 py-1.5 bg-slate-100 hover:bg-slate-200 rounded border">
            Прикрепить файлы…
          </button>
          {ev.oscillograms.length > 0 && (
            <ul className="mt-2 space-y-1">
              {ev.oscillograms.map((o) => (
                <li key={o.id} className="flex items-center gap-2 text-sm">
                  <FileText size={14} className="text-slate-400 shrink-0" />
                  <span className="truncate">{o.name}</span>
                  <button type="button" className="text-red-500 hover:text-red-700 ml-auto"
                    onClick={() => setEv((p) => ({ ...p, oscillograms: p.oscillograms.filter((x) => x.id !== o.id) }))}>
                    <Trash2 size={14} />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="flex justify-end gap-2 mt-5">
          <button onClick={onClose} className="px-4 py-2 rounded bg-slate-200 hover:bg-slate-300 text-sm font-medium">Отмена</button>
          <button onClick={() => onSave(ev)} className="px-4 py-2 rounded bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium">Сохранить событие</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Импорт из Excel ----------
function ImportModal({
  sheets, onConfirm, onClose,
}: { sheets: ParsedSheet[]; onConfirm: (eventsBySheet: Map<string, OutageEvent[]>) => void; onClose: () => void }) {
  // состояние выбора столбцов по листам
  const [colsState, setColsState] = useState<Map<string, ImportColumn[]>>(
    () => new Map(sheets.map((s) => [s.sheetName, s.columns.map((c) => ({ ...c }))]))
  );
  const [activeSheet, setActiveSheet] = useState(sheets[0]?.sheetName ?? '');

  const sortedSheets = useMemo(
    () => [...sheets].sort((a, b) => (a.year ?? 9999) - (b.year ?? 9999)),
    [sheets]
  );

  const toggleCol = (sheet: string, idx: number, include: boolean) => {
    setColsState((prev) => {
      const m = new Map(prev);
      const cols = (m.get(sheet) ?? []).map((c, i) => (i === idx ? { ...c, include } : c));
      m.set(sheet, cols);
      return m;
    });
  };

  const doImport = () => {
    const result = new Map<string, OutageEvent[]>();
    for (const s of sortedSheets) {
      const cols = colsState.get(s.sheetName) ?? s.columns;
      // ключи столбцов, реально присутствовавшие в исходном листе Excel
      const presentKeys = new Set(
        s.columns.filter((c) => c.targetKey && c.targetKey !== 'oscillograms').map((c) => c.targetKey as string)
      );
      // пересобираем события с учётом выбранных столбцов
      const events: OutageEvent[] = [];
      for (const ev of s.rows[0]) {
        const clone: OutageEvent = { ...EMPTY_EVENT, id: ev.id, oscillograms: [] };
        let hasData = false;
        cols.forEach((c) => {
          if (!c.include || !c.targetKey || c.targetKey === 'oscillograms') return;
          const val = (ev as any)[c.targetKey];
          if (val === '' || val == null || (c.targetKey === 'refusal' && val === false)) return;
          (clone as any)[c.targetKey] = val;
          hasData = true;
        });
        // пустые ячейки (по столбцам, которые были в исходной таблице) заполняем символом "-",
        // чтобы импортированные события не попадали в категорию "не оформленные"
        REQUIRED_FIELDS.forEach((f) => {
          if (presentKeys.has(f as string)) return; // столбец был в листе и остался пуст — это настоящий пробел
          const cur = (clone as any)[f];
          if (typeof cur === 'string' && cur.trim() === '') (clone as any)[f] = '-';
        });
        if (hasData) events.push(clone);
      }
      result.set(s.sheetName, events);
    }
    onConfirm(result);
  };

  const active = sortedSheets.find((s) => s.sheetName === activeSheet) ?? sortedSheets[0];
  const activeCols = active ? colsState.get(active.sheetName) ?? active.columns : [];
  const totalNew = sortedSheets.reduce((acc, s) => acc + (s.rows[0]?.length ?? 0), 0);

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={onClose}>
      <div className="bg-white rounded-lg shadow-xl p-6 max-w-2xl w-full mx-4 max-h-[85vh] overflow-y-auto"
        onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-3">
          <h2 className="text-lg font-bold">Импорт событий из Excel</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={20} /></button>
        </div>
        <p className="text-sm text-slate-600 mb-3">
          Найдено листов (годов): <b>{sortedSheets.length}</b>, событий для импорта: <b>{totalNew}</b>.
          События будут отсортированы по годам. Отметьте столбцы, которые нужно импортировать.
          Столбцы «Аварийные параметры на ПС №1/№2» удалены из формы с 2020 г. и по умолчанию не импортируются.
        </p>

        <div className="flex flex-wrap gap-1 mb-3">
          {sortedSheets.map((s) => (
            <button key={s.sheetName} onClick={() => setActiveSheet(s.sheetName)}
              className={`px-3 py-1.5 rounded text-sm font-medium border ${
                activeSheet === s.sheetName ? 'bg-blue-600 text-white border-blue-600' : 'bg-white hover:bg-slate-100'
              }`}>
              {s.year ? `${s.year}` : s.sheetName}
            </button>
          ))}
        </div>

        {active && (
          <div className="border rounded-md">
            <div className="bg-slate-50 px-3 py-2 text-sm font-semibold border-b">
              Лист «{active.sheetName}» {active.year ? `(${active.year} г.)` : ''} — {active.rows[0]?.length ?? 0} событий
            </div>
            <div className="max-h-64 overflow-y-auto divide-y">
              {activeCols.map((c, i) => (
                <label key={i} className="flex items-center gap-2 px-3 py-1.5 text-sm hover:bg-slate-50 cursor-pointer">
                  <input type="checkbox" checked={c.include} disabled={!c.targetKey}
                    onChange={(e) => toggleCol(active.sheetName, i, e.target.checked)}
                    className="w-4 h-4 accent-blue-600" />
                  <span className={c.targetKey ? '' : 'text-slate-400'}>{c.excelHeader}</span>
                  {!c.targetKey && (
                    <span className="ml-auto text-xs px-2 py-0.5 rounded bg-amber-100 text-amber-800">
                      не импортируется (нет в новой форме)
                    </span>
                  )}
                  {c.targetKey && (
                    <span className="ml-auto text-xs text-slate-400">
                      → {COLUMNS.find((col) => col.key === c.targetKey)?.title ?? (c.targetKey === 'refusal' ? 'Отказ' : '')}
                    </span>
                  )}
                </label>
              ))}
            </div>
          </div>
        )}

        <div className="flex justify-end gap-2 mt-4">
          <button onClick={onClose} className="px-4 py-2 rounded bg-slate-200 hover:bg-slate-300 text-sm font-medium">Отмена</button>
          <button onClick={doImport} className="px-4 py-2 rounded bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium">
            Импортировать
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Фильтр по значению столбца ----------
function ColumnFilterPopover({
  colKey, title, events, selected, onApply, onClose, anchorRect,
}: {
  colKey: keyof OutageEvent; title: string; events: OutageEvent[];
  selected: Set<string>; onApply: (vals: Set<string>) => void; onClose: () => void;
  anchorRect: DOMRect | null;
}) {
  const values = useMemo(() => {
    const set = new Set<string>();
    events.forEach((e) => {
      if (colKey === 'num') { set.add(String(e.num ?? '')); return; }
      const v = String((e as any)[colKey] ?? '').trim();
      if (v) set.add(v);
    });
    if (colKey === 'num') {
      return [...set].filter(Boolean).sort((a, b) => +a - +b);
    }
    return [...set].sort((a, b) => a.localeCompare(b, 'ru'));
  }, [events, colKey]);
  const [checked, setChecked] = useState<Set<string>>(new Set(selected));
  const [search, setSearch] = useState('');

  const shown = values.filter((v) => v.toLowerCase().includes(search.toLowerCase()));

  const left = anchorRect ? Math.max(8, Math.min(anchorRect.left, window.innerWidth - 280)) : 8;
  const top = anchorRect ? anchorRect.bottom + 6 : 8;

  return createPortal(
    <>
      <div className="fixed inset-0 z-[70]" onClick={onClose} />
      <div className="fixed z-[80] bg-white text-slate-800 border border-slate-300 rounded-md shadow-xl w-64 p-2"
        style={{ left, top, maxHeight: '60vh', display: 'flex', flexDirection: 'column' }}>
        <div className="text-xs font-semibold text-slate-600 mb-1">{title}</div>
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Поиск значения…" autoFocus
          className="w-full text-sm border border-slate-300 rounded px-2 py-1 mb-1" />
        <div className="overflow-y-auto" style={{ maxHeight: '35vh' }}>
          {shown.length === 0 && <div className="text-xs text-slate-400 px-1 py-2">Нет значений</div>}
          {shown.map((v) => (
            <label key={v} className="flex items-start gap-2 text-sm px-1 py-0.5 cursor-pointer hover:bg-slate-50">
              <input type="checkbox" className="mt-1 accent-blue-600" checked={checked.has(v)}
                onChange={(e) => {
                  setChecked((p) => { const n = new Set(p); e.target.checked ? n.add(v) : n.delete(v); return n; });
                }} />
              <span className="break-all">{v}</span>
            </label>
          ))}
        </div>
        <div className="flex justify-between mt-2">
          <button className="text-xs text-blue-600 hover:underline" onClick={() => setChecked(new Set())}>Сбросить</button>
          <button className="text-xs px-2 py-1 bg-blue-600 text-white rounded hover:bg-blue-700"
            onClick={() => onApply(checked)}>OK</button>
        </div>
      </div>
    </>,
    document.body
  );
}

// ---------- Основная таблица / приложение ----------
export default function App() {
  const [events, setEvents] = useState<OutageEvent[]>(loadEvents);
  const [savedSnapshot, setSavedSnapshot] = useState<string>(() => JSON.stringify(loadEvents()));
  const [editing, setEditing] = useState<OutageEvent | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [quickFilter, setQuickFilter] = useState<QuickFilter>('none');
  const [colFilters, setColFilters] = useState<Map<keyof OutageEvent, Set<string>>>(new Map());
  const [openFilterCol, setOpenFilterCol] = useState<{ key: keyof OutageEvent; rect: DOMRect } | null>(null);
  const [searchText, setSearchText] = useState('');
  const [activeYear, setActiveYear] = useState<number | 'all' | 'nodate'>('all');
  const [confirmClear, setConfirmClear] = useState(false);
  const [confirmExit, setConfirmExit] = useState(false);
  const [importSheets, setImportSheets] = useState<ParsedSheet[] | null>(null);
  const [hoveredOsc, setHoveredOsc] = useState<{ evId: string; x: number; y: number } | null>(null);
  const [activeTab, setActiveTab] = useState<'table' | 'analysis'>('table');
  const jsonInputRef = useRef<HTMLInputElement>(null);
  const oscHideTimer = useRef<number | null>(null);
  const oscPopupRef = useRef<HTMLDivElement | null>(null);
  // «прищепка» тулtipа: если курсор зажат (идёт выделение текста в тултипе) — не скрываем окно
  const oscPinned = useRef(false);

  useEffect(() => {
    const up = () => { oscPinned.current = false; };
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  const clearOscHide = () => {
    if (oscHideTimer.current) { window.clearTimeout(oscHideTimer.current); oscHideTimer.current = null; }
  };

  const scheduleOscHide = () => {
    clearOscHide();
    oscHideTimer.current = window.setTimeout(() => {
      if (oscPinned.current) { scheduleOscHide(); return; }
      setHoveredOsc(null);
    }, 600);
  };

  // скачивание всех осциллограмм ячейки (одна кнопка для всех файлов)
  // файлы упаковываются в один ZIP-архив — браузер не блокирует множественную загрузку
  const oscUrl = (o: Oscillogram) => encodeURI(`Осциллограммы/${o.fileName}`);

  const downloadAllOsc = async (list: Oscillogram[], eventId?: number) => {
    if (list.length === 1) {
      const a = document.createElement('a');
      a.href = oscUrl(list[0]);
      a.download = list[0].fileName || 'oscillogram';
      document.body.appendChild(a);
      a.click();
      a.remove();
      return;
    }
    try {
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();
      const usedNames = new Set<string>();
      await Promise.all(list.map(async (o, idx) => {
        let name = (o.fileName || `oscillogram_${idx + 1}`).replace(/[\\/:*?"<>|]/g, '_');
        let finalName = name;
        let dupIdx = 2;
        while (usedNames.has(finalName)) {
          const dot = name.lastIndexOf('.');
          finalName = dot > 0 ? `${name.slice(0, dot)}_${dupIdx}${name.slice(dot)}` : `${name}_${dupIdx}`;
          dupIdx++;
        }
        usedNames.add(finalName);
        const resp = await fetch(oscUrl(o));
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        zip.file(finalName, await resp.blob());
      }));
      const blob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Осциллограммы_событие_${eventId ?? ''}.zip`.replace('__', '_');
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch (e) {
      setToast('Не удалось упаковать файлы в архив (проверьте наличие файлов в папке «Осциллограммы» рядом со страницей)');
    }
  };
  const [lightbox, setLightbox] = useState<Oscillogram | null>(null);
  const [toast, setToast] = useState<string>('');
  const importInputRef = useRef<HTMLInputElement>(null);

  const dirty = JSON.stringify(events) !== savedSnapshot;

  const showToast = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 2500);
  };

  const save = useCallback(() => {
    const snap = JSON.stringify(events);
    try {
      localStorage.setItem(STORAGE_KEY, snap);
      setSavedSnapshot(snap);
      showToast('Данные сохранены');
    } catch (e) {
      showToast('Ошибка сохранения (возможно, превышен лимит хранилища из-за осциллограмм)');
    }
  }, [events]);

  // предупреждение при закрытии вкладки, если есть несохраненные изменения
  useEffect(() => {
    const handler = (e: BeforeUnloadEvent) => {
      if (dirty) { e.preventDefault(); e.returnValue = ''; }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [dirty]);

  // ---------- Сохранение всей базы событий в JSON-файл ----------
  const saveToJson = useCallback(() => {
    try {
      const blob = new Blob([JSON.stringify(events, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'outages.json';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      showToast('Файл outages.json сохранён');
    } catch (e) {
      showToast('Не удалось сохранить JSON');
    }
  }, [events]);

  // ---------- Загрузка базы событий из JSON-файла ----------
  const loadFromJson = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(String(reader.result));
        const list: OutageEvent[] = Array.isArray(data) ? data : (Array.isArray(data?.events) ? data.events : null);
        if (!list) { showToast('В файле не найден массив событий'); return; }
        const normalized = list.map((e, i) => ({
          ...EMPTY_EVENT,
          ...e,
          id: e.id ?? crypto.randomUUID(),
          num: e.num ?? i + 1,
          oscillograms: Array.isArray(e.oscillograms) ? e.oscillograms : [],
        }));
        setEvents(normalized);
        const snap = JSON.stringify(normalized);
        localStorage.setItem(STORAGE_KEY, snap);
        setSavedSnapshot(snap);
        showToast(`Загружено событий из JSON: ${normalized.length}`);
      } catch {
        showToast('Ошибка чтения JSON: файл повреждён или имеет неверный формат');
      }
    };
    reader.onerror = () => showToast('Не удалось прочитать файл');
    reader.readAsText(file);
  };

  // годы для вкладок (по умолчанию открывается самый последний год)
  const years = useMemo(() => {
    const set = new Set<number>();
    let hasNoDate = false;
    events.forEach((e) => {
      const y = getYear(e.date);
      if (y) set.add(y); else hasNoDate = true;
    });
    return { list: [...set].sort((a, b) => b - a), hasNoDate };
  }, [events]);

  const initialYearRef = useRef<number | null>(null);
  useEffect(() => {
    if (initialYearRef.current == null && years.list.length > 0) {
      initialYearRef.current = years.list[0];
      setActiveYear(years.list[0]);
    }
  }, [years]);

  useEffect(() => {
    // если активная вкладка года исчезла — вернуться на последнюю вкладку
    if (typeof activeYear === 'number' && !years.list.includes(activeYear)) {
      setActiveYear(years.list[0] ?? 'all');
    }
  }, [years, activeYear]);

  // Количество событий на вкладках лет С УЧЁТОМ быстрого фильтра,
  // но БЕЗ учёта активной вкладки года (activeYear) и поиска (searchText)
  const eventsForYearTabs = useMemo(() => {
    let list = events;
    if (quickFilter === 'refusal') list = list.filter((e) => e.refusal);
    else if (quickFilter === 'v330') list = list.filter((e) => e.voltage === '330');
    else if (quickFilter === 'v110') list = list.filter((e) => e.voltage === '110');
    else if (quickFilter === 'incomplete') list = list.filter((e) => isIncomplete(e));

    const counts = new Map<number, number>();
    let noDateCount = 0;
    list.forEach((e) => {
      const y = getYear(e.date);
      if (y) counts.set(y, (counts.get(y) || 0) + 1);
      else noDateCount++;
    });
    return { counts, noDateCount };
  }, [events, quickFilter]);

  const filtered = useMemo(() => {
    let list = events;
    // вкладка по году
    if (activeYear === 'nodate') list = list.filter((e) => !getYear(e.date));
    else if (typeof activeYear === 'number') list = list.filter((e) => getYear(e.date) === activeYear);

    if (quickFilter === 'refusal') list = list.filter((e) => e.refusal);
    else if (quickFilter === 'v330') list = list.filter((e) => e.voltage === '330');
    else if (quickFilter === 'v110') list = list.filter((e) => e.voltage === '110');
    else if (quickFilter === 'incomplete') list = list.filter((e) => isIncomplete(e));

    for (const [key, vals] of colFilters) {
      if (vals.size === 0) continue;
      list = list.filter((e) => {
        const v = key === 'num' ? String(e.num ?? '') : String((e as any)[key] ?? '').trim();
        return vals.has(v);
      });
    }

    // полнотекстовый поиск по всем столбцам
    const q = searchText.trim().toLowerCase();
    if (q) {
      list = list.filter((e) => {
        const parts: string[] = [];
        COLUMNS.forEach((c) => {
          if (c.key === 'oscillograms') {
            e.oscillograms.forEach((o) => parts.push(o.name));
          } else {
            parts.push(String((e as any)[c.key] ?? ''));
          }
        });
        parts.push(e.refusal ? 'отказ отказ' : '');
        return parts.join('\n').toLowerCase().includes(q);
      });
    }

    // сортировка: новые события сверху (по убыванию даты)
    return [...list].sort((a, b) => (parseDateStr(b.date)?.getTime() ?? 0) - (parseDateStr(a.date)?.getTime() ?? 0));
  }, [events, quickFilter, colFilters, searchText, activeYear]);

  const handleSaveEvent = (ev: OutageEvent) => {
    setEvents((prev) => {
      const idx = prev.findIndex((e) => e.id === ev.id);
      if (idx >= 0) { const n = [...prev]; n[idx] = ev; return n; }
      const maxNum = prev.reduce((m, e) => Math.max(m, e.num ?? 0), 0);
      return [...prev, { ...ev, num: ev.num ?? maxNum + 1 }];
    });
    setShowForm(false);
    setEditing(null);
  };

  const deleteEvent = (id: string) => {
    setEvents((prev) => prev.filter((e) => e.id !== id));
  };

  const handleImportFile = async (file: File) => {
    const buf = await file.arrayBuffer();
    try {
      const sheets = parseWorkbook(buf);
      if (!sheets.length) { showToast('В файле не найдено листов с данными'); return; }
      setImportSheets(sheets);
    } catch (e) {
      showToast('Не удалось прочитать файл Excel');
    }
  };

  const confirmImport = (bySheet: Map<string, OutageEvent[]>) => {
    const incoming: OutageEvent[] = [];
    // защита от дубликатов по дате (среди импортируемых и существующих событий)
    const seenDates = new Set<string>();
    let skipped = 0;
    // сортировка листов по году при импорте
    const ordered = [...bySheet.entries()].sort((a, b) => {
      const ya = getYear(a[1][0]?.date ?? '') ?? a[0].match(/(19|20)\d{2}/)?.[0] as any;
      const yb = getYear(b[1][0]?.date ?? '') ?? b[0].match(/(19|20)\d{2}/)?.[0] as any;
      return Number(ya ?? 0) - Number(yb ?? 0);
    });
    ordered.forEach(([, list]) => {
      for (const e of list) {
        const dateKey = (e.date ?? '').trim();
        if (dateKey && (events.some(ev => ev.date === dateKey) || seenDates.has(dateKey))) {
          skipped++;
          continue; // событие с такой датой уже есть — пропускаем
        }
        seenDates.add(dateKey);
        incoming.push(e);
      }
    });
    setEvents((prev) => {
      const maxNum = prev.reduce((m, e) => Math.max(m, e.num ?? 0), 0);
      let n = maxNum;
      const stamped = incoming.map((e) => ({ ...e, num: e.num ?? ++n }));
      return [...prev, ...stamped];
    });
    setImportSheets(null);
    if (incoming.length) {
      // открыть вкладку последнего импортированного года
      const lastYear = [...incoming].reverse().map((e) => getYear(e.date)).find((y) => y != null);
      if (lastYear) setActiveYear(lastYear);
    }
    showToast(`Импортировано событий: ${incoming.length}${skipped > 0 ? ` (дубликаты пропущены: ${skipped})` : ''}`);
  };

  const exitWithoutSave = () => {
    try {
      const restored = JSON.parse(savedSnapshot);
      setEvents(restored);
    } catch { /* ignore */ }
    window.close();
    // если браузер запретил close — покажем подсказку
    showToast('Закройте вкладку вручную (браузер блокирует автоматическое закрытие)');
  };

  const hoveredEvent = hoveredOsc ? events.find((e) => e.id === hoveredOsc.evId) : null;

  const btnCls = 'inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm font-medium transition-colors shadow-sm';

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800">
      <header className="bg-slate-800 text-white px-4 py-3 shadow sticky top-0 z-30">
        <div className="flex flex-wrap items-center gap-3 max-w-[1800px] mx-auto">
          <h1 className="text-lg font-bold mr-auto">Программа отключений <span className="font-normal text-slate-300 text-sm">— сеть 110 / 330 кВ</span></h1>
          <div className="flex flex-wrap gap-2">
            <button className={`${btnCls} bg-green-600 hover:bg-green-700 text-white`} onClick={() => { setEditing(null); setShowForm(true); }}>
              <Plus size={16} /> Добавить событие
            </button>
            <button className={`${btnCls} bg-blue-600 hover:bg-blue-700 text-white`} onClick={() => importInputRef.current?.click()}>
              <Upload size={16} /> Импорт событий
            </button>
            <input ref={importInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) handleImportFile(f); e.target.value = ''; }} />
            <button className={`${btnCls} bg-blue-600 hover:bg-blue-700 text-white`} onClick={() => exportToExcel(filtered.length ? filtered : events)}>
              <Download size={16} /> Экспорт событий
            </button>
            <button className={`${btnCls} ${dirty ? 'bg-amber-500 hover:bg-amber-600 text-white' : 'bg-slate-500 text-slate-200'}`} onClick={save}>
              <Save size={16} /> Сохранить {dirty && <span className="w-2 h-2 rounded-full bg-white inline-block" />}
            </button>
            <button className={`${btnCls} bg-emerald-600 hover:bg-emerald-700 text-white`} onClick={saveToJson} title="Выгрузить все события в файл outages.json">
              <FileJson size={16} /> Сохранить в JSON
            </button>
            <button className={`${btnCls} bg-emerald-700 hover:bg-emerald-800 text-white`} onClick={() => jsonInputRef.current?.click()} title="Загрузить события из файла outages.json">
              <FolderOpen size={16} /> Загрузить из JSON
            </button>
            <input ref={jsonInputRef} type="file" accept=".json" className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) loadFromJson(f); e.target.value = ''; }} />
            <button className={`${btnCls} bg-red-600 hover:bg-red-700 text-white`} onClick={() => setConfirmClear(true)}>
              <Trash2 size={16} /> Очистить таблицу
            </button>
            <button className={`${btnCls} bg-slate-600 hover:bg-slate-500 text-white`} onClick={() => setConfirmExit(true)}>
              <LogOut size={16} /> Выйти
            </button>
          </div>
        </div>
      </header>

      {/* Быстрые фильтры + поиск */}
      <div className="max-w-[1800px] mx-auto px-4 pt-3 flex flex-wrap items-center gap-2">
        <Filter size={16} className="text-slate-500" />
        <span className="text-sm text-slate-600 font-medium">Быстрые фильтры:</span>
        <button onClick={() => setQuickFilter('none')}
          className={`px-3 py-1 rounded-full text-sm border ${quickFilter === 'none' ? 'bg-slate-800 text-white border-slate-800' : 'bg-white hover:bg-slate-50'}`}>
          Все ({events.length})
        </button>
        {QUICK_FILTERS.map((f) => {
          const count = (() => {
            if (f.id === 'refusal') return events.filter((e) => e.refusal).length;
            if (f.id === 'v330') return events.filter((e) => e.voltage === '330').length;
            if (f.id === 'v110') return events.filter((e) => e.voltage === '110').length;
            return events.filter((e) => isIncomplete(e)).length;
          })();
          const activeQ = quickFilter === f.id;
          return (
            <button key={f.id} onClick={() => setQuickFilter(activeQ ? 'none' : f.id)}
              className={`px-3 py-1 rounded-full text-sm border ${
                activeQ ? (f.id === 'refusal' ? 'bg-red-600 text-white border-red-600' : 'bg-blue-600 text-white border-blue-600')
                : 'bg-white hover:bg-slate-50'
              }`}>
              {f.label} ({count})
            </button>
          );
        })}
        {colFilters.size > 0 && (
          <button className="text-sm text-blue-700 underline ml-2"
            onClick={() => { setColFilters(new Map()); }}>
            Сбросить фильтры столбцов
          </button>
        )}
        <div className="relative ml-auto">
          <input
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            placeholder="Поиск по тексту (дата, линия, причина, примечания и т.д.)…"
            className="w-[340px] max-w-full text-sm border border-slate-300 rounded-md px-3 py-1.5 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          {searchText && (
            <button className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700"
              onClick={() => setSearchText('')} title="Очистить поиск">
              <X size={14} />
            </button>
          )}
        </div>
      </div>

      {/* Переключатель вкладок: Таблица / Анализ */}
      <div className="max-w-[1800px] mx-auto px-4 pt-3 flex flex-wrap items-center gap-1">
        <button onClick={() => setActiveTab('table')}
          className={`inline-flex items-center gap-1.5 px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
            activeTab === 'table' ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
          }`}>
          <TableIcon size={14} /> Таблица
        </button>
        <button onClick={() => setActiveTab('analysis')}
          className={`inline-flex items-center gap-1.5 px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
            activeTab === 'analysis' ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
          }`}>
          <BarChart3 size={14} /> Анализ
        </button>
      </div>

      {activeTab === 'analysis' ? (
        <main className="max-w-[1800px] mx-auto p-4">
          {/* Анализ строится по отфильтрованным событиям: учтены вкладки лет и быстрый поиск */}
          <AnalysisTab events={filtered} />
        </main>
      ) : (
      <>
      {/* Вкладки по годам */}
      <div className="max-w-[1800px] mx-auto px-4 pt-3 flex flex-wrap items-center gap-1 border-b border-slate-200 pb-0">
        <button onClick={() => setActiveYear('all')}
          className={`px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
            activeYear === 'all' ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
          }`}>
          Все годы ({events.length})
        </button>
        {years.list.map((y) => {
          const cnt = eventsForYearTabs.counts.get(y) || 0;
          return (
            <button key={y} onClick={() => setActiveYear(y)}
              className={`px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
                activeYear === y ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
              }`}>
              {y} г. ({cnt})
            </button>
          );
        })}
        {years.hasNoDate && (
          <button onClick={() => setActiveYear('nodate')}
            className={`px-4 py-1.5 rounded-t-md text-sm font-semibold border border-b-0 ${
              activeYear === 'nodate' ? 'bg-white border-slate-300 text-blue-700 -mb-px' : 'bg-slate-200/70 text-slate-600 hover:bg-slate-100 border-transparent'
            }`}>
            Без даты ({eventsForYearTabs.noDateCount})
          </button>
        )}
      </div>

      {/* Таблица */}
      <main className="max-w-[1800px] mx-auto p-4">
        <div className="bg-white rounded-lg shadow overflow-auto border" style={{ maxHeight: 'calc(100vh - 230px)' }}>
          <table className="border-collapse text-sm w-full" style={{ minWidth: 1750, tableLayout: 'fixed' }}>
            <thead className="sticky top-0 z-20">
              <tr className="bg-slate-700 text-white">
                <th className="border border-slate-600 px-2 py-2 w-16 relative">
                  <div className="flex items-center gap-1">
                    <span className="font-semibold">№</span>
                    <button className={`ml-auto shrink-0 ${colFilters.get('num')?.size ? 'text-amber-300' : 'text-slate-300 hover:text-white'}`}
                      onClick={(ev) => setOpenFilterCol(openFilterCol?.key === 'num' ? null : { key: 'num', rect: (ev.currentTarget as HTMLElement).getBoundingClientRect() })}
                      title="Фильтр по номеру события">
                      <Filter size={13} />
                    </button>
                  </div>
                  {openFilterCol?.key === 'num' && (
                    <ColumnFilterPopover
                      colKey={'num' as keyof OutageEvent} title="Номер события" events={events}
                      selected={colFilters.get('num' as keyof OutageEvent) ?? new Set()}
                      onApply={(vals) => {
                        setColFilters((prev) => {
                          const m = new Map(prev);
                          if (vals.size === 0) m.delete('num' as keyof OutageEvent); else m.set('num' as keyof OutageEvent, vals);
                          return m;
                        });
                        setOpenFilterCol(null);
                      }}
                      onClose={() => setOpenFilterCol(null)}
                      anchorRect={openFilterCol.rect}
                    />
                  )}
                </th>
                {COLUMNS.map((col) => {
                  const hasFilter = colFilters.get(col.key)?.size;
                  return (
                    <th key={String(col.key)} className="border border-slate-600 px-2 py-2 text-left relative"
                      style={{ width: col.width, minWidth: col.width }}>
                      <div className="flex items-center gap-1">
                        <span className="font-semibold leading-tight">{col.title}</span>
                        {col.key !== 'oscillograms' && (
                          <button className={`ml-auto shrink-0 ${hasFilter ? 'text-amber-300' : 'text-slate-300 hover:text-white'}`}
                            onClick={(ev) => setOpenFilterCol(openFilterCol?.key === col.key ? null : { key: col.key, rect: (ev.currentTarget as HTMLElement).getBoundingClientRect() })}
                            title="Фильтр по столбцу">
                            <Filter size={13} />
                          </button>
                        )}
                      </div>
                      {openFilterCol?.key === col.key && (
                        <ColumnFilterPopover
                          colKey={col.key} title={col.title} events={events}
                          selected={colFilters.get(col.key) ?? new Set()}
                          onApply={(vals) => {
                            setColFilters((prev) => {
                              const m = new Map(prev);
                              if (vals.size === 0) m.delete(col.key); else m.set(col.key, vals);
                              return m;
                            });
                            setOpenFilterCol(null);
                          }}
                          onClose={() => setOpenFilterCol(null)}
                          anchorRect={openFilterCol.rect}
                        />
                      )}
                    </th>
                  );
                })}
                <th className="border border-slate-600 px-2 py-2 w-20">Действия</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={COLUMNS.length + 2} className="text-center py-10 text-slate-400">
                    {events.length === 0
                      ? 'Нет событий. Нажмите «Добавить событие» или «Импорт событий».'
                      : 'По текущим фильтрам событий не найдено.'}
                  </td>
                </tr>
              )}
              {filtered.map((e, i) => {
                const incomplete = isIncomplete(e);
                return (
                  <tr key={e.id}
                    className={e.refusal ? 'bg-red-600 text-white' : i % 2 ? 'bg-slate-50' : 'bg-white'}>
                    <td className={`border px-2 py-1 text-center align-top ${e.refusal ? 'border-red-500' : 'border-slate-200'}`}>{e.num ?? i + 1}</td>
                    {COLUMNS.map((col) => {
                      if (col.key === 'oscillograms') {
                        return (
                          <td key="osc" className={`border px-2 py-1 align-top ${e.refusal ? 'border-red-500' : 'border-slate-200'}`}
                            onMouseEnter={(ev) => {
                              if (e.oscillograms.length === 0) return;
                              clearOscHide();
                              setHoveredOsc({ evId: e.id, x: ev.clientX, y: ev.clientY });
                            }}
                            onMouseLeave={scheduleOscHide}>
                            {e.oscillograms.length > 0 ? (
                              <div className={`flex flex-col gap-0.5 cursor-help ${e.refusal ? 'text-white' : ''}`}>
                                {/* Относительные ссылки на файлы осциллограмм (папка «Осциллограммы» рядом со страницей) */}
                                {e.oscillograms.map((o) => (
                                  <a key={o.id}
                                    href={`Осциллограммы/${o.fileName}`}
                                    target="_blank"
                                    rel="noreferrer"
                                    title={o.name}
                                    onClick={(ev) => { ev.stopPropagation(); clearOscHide(); setHoveredOsc(null); }}
                                    className={`truncate text-xs hover:underline ${e.refusal ? 'text-white underline decoration-red-300' : 'text-blue-600'}`}>
                                    {o.fileName}
                                  </a>
                                ))}
                              </div>
                            ) : (
                              <span className={e.refusal ? 'text-red-200' : 'text-slate-300'}>—</span>
                            )}
                          </td>
                        );
                      }
                      const val = String((e as any)[col.key] ?? '');
                      const empty = val.trim() === '';
                      return (
                        <td key={String(col.key)}
                          className={`border px-2 py-1 align-top break-words ${e.refusal ? 'border-red-500' : 'border-slate-200'}`}
                          style={{ whiteSpace: 'pre-wrap' }}>
                          {empty ? (
                            <span className={e.refusal ? 'text-red-200 italic' : 'text-slate-300 italic'}>пусто</span>
                          ) : col.key === 'voltage' ? `${val} кВ` : val}
                        </td>
                      );
                    })}
                    <td className={`border px-1 py-1 text-center whitespace-nowrap align-top ${e.refusal ? 'border-red-500' : 'border-slate-200'}`}>
                      <button className="text-blue-600 hover:text-blue-800 px-1" title="Редактировать"
                        onClick={() => { setEditing(e); setShowForm(true); }}>✎</button>
                      <button className="text-red-500 hover:text-red-700 px-1" title="Удалить"
                        onClick={() => deleteEvent(e.id)}>🗑</button>
                      {incomplete && !e.refusal && (
                        <span title="Событие недооформлено" className="text-amber-500 px-0.5">◔</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="text-xs text-slate-500 mt-2">
          Показано: {filtered.length} из {events.length}. Красным цветом выделены события с отказом. ◔ — недооформленное событие (есть пустые ячейки).
        </div>
      </main>
      </>
      )}

      {/* Тултип с осциллограммами при наведении */}
      {hoveredOsc && hoveredEvent && hoveredEvent.oscillograms.length > 0 && createPortal(
        <div ref={oscPopupRef} className="fixed z-[90] bg-white border border-slate-300 rounded-lg shadow-xl p-3 w-80"
          style={{
            left: Math.min(hoveredOsc.x + 12, window.innerWidth - 340),
            top: Math.min(hoveredOsc.y + 12, window.innerHeight - 360),
          }}
          onMouseEnter={clearOscHide}
          onMouseLeave={scheduleOscHide}
          onMouseDown={() => { oscPinned.current = true; clearOscHide(); }}
          onClick={(ev) => ev.stopPropagation()}>
          <div className="flex items-center justify-between mb-2">
            <div className="text-xs font-bold text-slate-600 truncate">Осциллограммы: {hoveredEvent.line || 'без линии'}</div>
            <button title="Скачать все файлы ячейки"
              className="shrink-0 inline-flex items-center gap-1 text-xs px-2 py-1 rounded bg-blue-600 hover:bg-blue-700 text-white"
              onMouseDown={() => { oscPinned.current = true; }}
              onClick={(ev) => { ev.stopPropagation(); downloadAllOsc(hoveredEvent.oscillograms, hoveredEvent.num); }}>
              <Download size={12} /> Скачать все ({hoveredEvent.oscillograms.length})
            </button>
          </div>
          <div className="space-y-2 max-h-72 overflow-y-auto">
            {hoveredEvent.oscillograms.map((o) => (
              <div key={o.id} className="border rounded p-1 cursor-pointer hover:bg-slate-50" onClick={() => { setLightbox(o); setHoveredOsc(null); }}>
                <div className="flex items-center gap-1 mb-1">
                  <div className="text-xs truncate flex-1">{o.name}</div>
                  <a href={oscUrl(o)} download={o.name} title="Скачать файл"
                    className="shrink-0 text-blue-600 hover:text-blue-800"
                    onMouseDown={() => { oscPinned.current = true; }}
                    onClick={(ev) => ev.stopPropagation()}><Download size={13} /></a>
                </div>
                {o.type.startsWith('image/') ? (
                  <img src={oscUrl(o)} alt={o.name} className="w-full max-h-32 object-contain bg-slate-50 rounded" />
                ) : (
                  <div className="text-xs text-slate-500 flex items-center gap-1"><FileText size={14} /> {o.type || 'файл'} — нажмите для открытия</div>
                )}
              </div>
            ))}
          </div>
        </div>,
        document.body
      )}

      {/* Лайтбокс просмотра осциллограммы */}
      {lightbox && createPortal(
        <div className="fixed inset-0 bg-black/80 z-[100] flex items-center justify-center" onClick={() => setLightbox(null)}>
          <div className="max-w-[90vw] max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
            <div className="text-white text-sm mb-2 flex items-center gap-3">
              <span className="truncate">{lightbox.name}</span>
              <a href={oscUrl(lightbox)} download={lightbox.name}
                className="text-blue-300 underline shrink-0">Скачать</a>
              <button className="ml-auto text-white hover:text-red-300" onClick={() => setLightbox(null)}><X size={20} /></button>
            </div>
            {lightbox.type.startsWith('image/') ? (
              <img src={oscUrl(lightbox)} alt={lightbox.name} className="max-w-[90vw] max-h-[82vh] object-contain rounded bg-white" />
            ) : lightbox.type === 'application/pdf' ? (
              <iframe src={oscUrl(lightbox)} title={lightbox.name} className="w-[80vw] h-[80vh] bg-white rounded" />
            ) : (
              <div className="bg-white p-6 rounded text-sm">Предпросмотр недоступен. <a className="text-blue-600 underline" href={oscUrl(lightbox)} download={lightbox.name}>Скачать файл</a></div>
            )}
          </div>
        </div>,
        document.body
      )}

      {/* Модальные окна */}
      {showForm && <EventForm initial={editing} onSave={handleSaveEvent} onClose={() => { setShowForm(false); setEditing(null); }} />}

      {importSheets && <ImportModal sheets={importSheets} onConfirm={confirmImport} onClose={() => setImportSheets(null)} />}

      {confirmClear && (
        <ConfirmModal title="Вы уверены, что хотите очистить таблицу?"
          onClose={() => setConfirmClear(false)}
          buttons={[
            { label: 'Да', danger: true, onClick: () => { setEvents([]); setConfirmClear(false); showToast('Таблица очищена (не забудьте «Сохранить»)'); } },
            { label: 'Нет', onClick: () => setConfirmClear(false) },
          ]} />
      )}

      {confirmExit && (
        <ConfirmModal title={dirty ? 'В таблице есть несохраненные изменения. Выйти?' : 'Выйти из программы?'}
          onClose={() => setConfirmExit(false)}
          buttons={[
            { label: 'Сохранить и выйти', onClick: () => { save(); setConfirmExit(false); exitWithoutSave(); } },
            { label: 'Выйти без сохранения', danger: true, onClick: () => { setConfirmExit(false); exitWithoutSave(); } },
            { label: 'Отмена', onClick: () => setConfirmExit(false) },
          ]} />
      )}

      {toast && (
        <div className="fixed bottom-4 right-4 z-50 bg-slate-800 text-white px-4 py-2 rounded-lg shadow-lg text-sm flex items-center gap-2">
          <AlertTriangle size={15} className="hidden" />{toast}
        </div>
      )}
    </div>
  );
}
