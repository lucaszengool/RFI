'use client';

import { useState, useCallback, useRef } from 'react';

interface CompareResult {
  onlyInFile1: string[];
  onlyInFile2: string[];
  inBoth: string[];
}

interface FileSlot {
  name: string;
  headers: string[];
  rows: string[][];
  selectedColumn: string;
}

// Strip UTF-8 BOM if present
function stripBOM(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

// Robust CSV parser: handles quoted fields, commas inside quotes
function parseCSV(raw: string): string[][] {
  const text = stripBOM(raw);
  const rows: string[][] = [];
  const lines = text.split(/\r?\n/);
  for (const line of lines) {
    if (!line.trim()) continue;
    const cols: string[] = [];
    let cur = '';
    let inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuote && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuote = !inQuote;
        }
      } else if (ch === ',' && !inQuote) {
        cols.push(cur.trim());
        cur = '';
      } else {
        cur += ch;
      }
    }
    cols.push(cur.trim());
    rows.push(cols);
  }
  return rows;
}

// Get unique non-empty values from a named column
function getColumnValues(rows: string[][], headerName: string): string[] {
  if (rows.length === 0) return [];
  const headers = rows[0];
  const idx = headers.findIndex(h => h === headerName);
  if (idx === -1) return [];
  const values = new Set<string>();
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][idx]?.trim();
    if (val && val !== 'null') values.add(val);
  }
  return Array.from(values).sort();
}

// Prefer these column names when auto-selecting
const PREFERRED_COLS_FILE1 = ['物理园区', '园区'];
const PREFERRED_COLS_FILE2 = ['园区', '物理园区'];

function pickDefaultColumn(headers: string[], preferred: string[]): string {
  for (const p of preferred) {
    if (headers.includes(p)) return p;
  }
  return headers[0] ?? '';
}

// ─── File upload zone ───────────────────────────────────────────────────────

function FileUploadZone({
  label,
  slot,
  preferredCols,
  onFile,
  onColumnChange,
  dragHint,
  accent,
}: {
  label: string;
  slot: FileSlot | null;
  preferredCols: string[];
  onFile: (file: File) => void;
  onColumnChange: (col: string) => void;
  dragHint: string;
  accent: 'blue' | 'violet';
}) {
  const [dragActive, setDragActive] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragActive(false);
      const file = e.dataTransfer.files?.[0];
      if (file) onFile(file);
    },
    [onFile],
  );

  const borderActive = accent === 'blue' ? 'border-blue-400 bg-blue-950/20' : 'border-violet-400 bg-violet-950/20';
  const borderIdle = 'border-gray-600 hover:border-gray-400 bg-gray-900/60';

  return (
    <div className="flex flex-col gap-3 h-full">
      <div className="text-sm font-semibold text-gray-200">{label}</div>

      <div
        className={`relative border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all ${
          dragActive ? borderActive : borderIdle
        }`}
        onClick={() => inputRef.current?.click()}
        onDragOver={e => { e.preventDefault(); setDragActive(true); }}
        onDragLeave={() => setDragActive(false)}
        onDrop={handleDrop}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".csv"
          className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }}
        />
        {slot ? (
          <div className="space-y-1.5">
            <div className="text-green-400 font-medium text-sm break-all">{slot.name}</div>
            <div className="text-gray-500 text-xs">
              {slot.rows.length - 1} 行 · {slot.headers.length} 列
            </div>
          </div>
        ) : (
          <div className="space-y-2 py-2">
            <div className="text-3xl text-gray-600">↑</div>
            <div className="text-gray-400 text-sm">点击或拖拽 CSV 文件</div>
            <div className="text-gray-600 text-xs">{dragHint}</div>
          </div>
        )}
      </div>

      {slot && (
        <div className="space-y-2">
          <label className="text-xs text-gray-400">选择要对比的列：</label>
          <select
            value={slot.selectedColumn}
            onChange={e => onColumnChange(e.target.value)}
            className="w-full bg-gray-800 border border-gray-600 rounded-lg px-3 py-2 text-sm text-gray-100 focus:outline-none focus:border-blue-400"
          >
            {slot.headers.map(h => (
              <option key={h} value={h}>
                {h}{preferredCols.includes(h) ? ' ★' : ''}
              </option>
            ))}
          </select>
          <div className="text-xs text-gray-500">
            已选「{slot.selectedColumn}」列，共{' '}
            <span className="text-gray-300 font-medium">
              {getColumnValues(slot.rows, slot.selectedColumn).length}
            </span>{' '}
            个唯一值
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Result card ─────────────────────────────────────────────────────────────

const COLOR_MAP = {
  amber: {
    header: 'text-amber-400',
    badge: 'bg-amber-900/30 text-amber-200 hover:bg-amber-900/50',
    border: 'border-amber-800/40',
    count: 'text-amber-300',
  },
  blue: {
    header: 'text-blue-400',
    badge: 'bg-blue-900/30 text-blue-200 hover:bg-blue-900/50',
    border: 'border-blue-800/40',
    count: 'text-blue-300',
  },
  green: {
    header: 'text-green-400',
    badge: 'bg-green-900/30 text-green-200 hover:bg-green-900/50',
    border: 'border-green-800/40',
    count: 'text-green-300',
  },
} as const;

function ResultCard({
  title,
  subtitle,
  items,
  color,
  emptyText,
}: {
  title: string;
  subtitle: string;
  items: string[];
  color: keyof typeof COLOR_MAP;
  emptyText: string;
}) {
  const c = COLOR_MAP[color];

  return (
    <div className={`bg-gray-900 rounded-2xl border ${c.border} p-4 flex flex-col gap-3`}>
      <div className="flex items-start justify-between gap-2">
        <div>
          <div className={`font-semibold text-sm ${c.header}`}>{title}</div>
          {subtitle && <div className="text-gray-500 text-xs mt-0.5">{subtitle}</div>}
        </div>
        <span className={`text-xl font-bold ${c.count} shrink-0`}>{items.length}</span>
      </div>

      <div className="flex flex-col gap-1 overflow-y-auto" style={{ maxHeight: '480px' }}>
        {items.length === 0 ? (
          <div className="text-gray-600 text-sm italic py-4 text-center">{emptyText}</div>
        ) : (
          items.map((item, i) => (
            <div
              key={item}
              className={`flex items-center gap-2 px-2.5 py-1.5 rounded-lg text-sm cursor-default transition-colors ${c.badge}`}
            >
              <span className="text-gray-600 text-xs w-6 text-right shrink-0 tabular-nums">{i + 1}</span>
              <span className="flex-1 min-w-0 truncate">{item}</span>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ─── Main page ───────────────────────────────────────────────────────────────

export default function ComparePage() {
  const [file1, setFile1] = useState<FileSlot | null>(null);
  const [file2, setFile2] = useState<FileSlot | null>(null);
  const [result, setResult] = useState<CompareResult | null>(null);

  const loadFile = useCallback(
    async (file: File, preferredCols: string[], setSlot: (s: FileSlot) => void) => {
      const raw = await file.text();
      const rows = parseCSV(raw);
      const headers = rows[0] ?? [];
      const selectedColumn = pickDefaultColumn(headers, preferredCols);
      setSlot({ name: file.name, headers, rows, selectedColumn });
      setResult(null);
    },
    [],
  );

  const handleFile1 = useCallback(
    (f: File) => loadFile(f, PREFERRED_COLS_FILE1, s => setFile1(s)),
    [loadFile],
  );
  const handleFile2 = useCallback(
    (f: File) => loadFile(f, PREFERRED_COLS_FILE2, s => setFile2(s)),
    [loadFile],
  );

  const handleColumn1 = useCallback((col: string) => {
    setFile1(prev => (prev ? { ...prev, selectedColumn: col } : null));
    setResult(null);
  }, []);
  const handleColumn2 = useCallback((col: string) => {
    setFile2(prev => (prev ? { ...prev, selectedColumn: col } : null));
    setResult(null);
  }, []);

  const handleCompare = useCallback(() => {
    if (!file1 || !file2) return;
    const set1 = new Set(getColumnValues(file1.rows, file1.selectedColumn));
    const set2 = new Set(getColumnValues(file2.rows, file2.selectedColumn));
    setResult({
      onlyInFile1: [...set1].filter(v => !set2.has(v)).sort(),
      onlyInFile2: [...set2].filter(v => !set1.has(v)).sort(),
      inBoth: [...set1].filter(v => set2.has(v)).sort(),
    });
  }, [file1, file2]);

  const canCompare = !!(file1?.selectedColumn && file2?.selectedColumn);

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100">
      {/* Top bar */}
      <header className="border-b border-gray-800 bg-gray-900/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-5 py-3 flex items-center gap-4">
          <a href="/" className="text-gray-500 hover:text-gray-300 text-sm transition-colors">← 返回</a>
          <div>
            <h1 className="text-base font-semibold">CSV 园区对比工具</h1>
            <p className="text-xs text-gray-500">上传两个 CSV，选列，自动找出差异</p>
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-5 py-6 space-y-6">
        {/* Upload cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="bg-gray-900 rounded-2xl p-5 border border-gray-800">
            <FileUploadZone
              label="文件 1 — 建设明细"
              slot={file1}
              preferredCols={PREFERRED_COLS_FILE1}
              onFile={handleFile1}
              onColumnChange={handleColumn1}
              dragHint='含「物理园区」列'
              accent="blue"
            />
          </div>
          <div className="bg-gray-900 rounded-2xl p-5 border border-gray-800">
            <FileUploadZone
              label="文件 2 — 电费预测"
              slot={file2}
              preferredCols={PREFERRED_COLS_FILE2}
              onFile={handleFile2}
              onColumnChange={handleColumn2}
              dragHint='含「园区」列'
              accent="violet"
            />
          </div>
        </div>

        {/* Compare button */}
        <button
          onClick={handleCompare}
          disabled={!canCompare}
          className={`w-full py-3 rounded-xl font-semibold text-sm transition-all ${
            canCompare
              ? 'bg-blue-600 hover:bg-blue-500 text-white shadow-lg shadow-blue-900/30'
              : 'bg-gray-800 text-gray-500 cursor-not-allowed'
          }`}
        >
          {canCompare ? '开始比对' : '请先上传两个文件'}
        </button>

        {/* Results */}
        {result && (
          <>
            <div className="text-xs text-gray-500 text-center">
              文件1「{file1?.selectedColumn}」vs 文件2「{file2?.selectedColumn}」
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <ResultCard
                title="仅在文件 1 中"
                subtitle={`列：${file1?.selectedColumn}`}
                items={result.onlyInFile1}
                color="amber"
                emptyText="无差异"
              />
              <ResultCard
                title="仅在文件 2 中"
                subtitle={`列：${file2?.selectedColumn}`}
                items={result.onlyInFile2}
                color="blue"
                emptyText="无差异"
              />
              <ResultCard
                title="两个文件都有"
                subtitle=""
                items={result.inBoth}
                color="green"
                emptyText="无共同项"
              />
            </div>
          </>
        )}
      </div>
    </div>
  );
}
