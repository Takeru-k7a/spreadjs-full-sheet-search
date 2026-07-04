'use client';

/**
 * このファイルは SpreadJS Workbook を実際に走査する検索エンジンです。
 * 全シート全件検索、アクティブセル起点の「次を検索」、
 * 検索結果からのセルジャンプ処理をここに集約しています。
 */

import * as GC from '@grapecity/spread-sheets';
import type { SearchHit, SearchOptions } from './searchStore';
import { matchesSearch, toA1Address } from './textNormalize';

export type SearchAllResult = {
  results: SearchHit[];
  truncated: boolean;
  message: string | null;
};

export type FindNextResult = {
  hit: SearchHit | null;
  message: string | null;
};

export type JumpResult = {
  ok: boolean;
  message: string | null;
};

export type SearchRuntimeOptions = {
  /** true の場合、検索時の Workbook/Sheet/走査範囲を console に出します。 */
  debug?: boolean;
  /** usedRange が空のとき、sheet の行列数を使って fallback 走査します。 */
  fallbackToSheetRange?: boolean;
  /** fallback 走査時の最大行数です。巨大シートで固まることを避けます。 */
  fallbackRowLimit?: number;
  /** fallback 走査時の最大列数です。巨大シートで固まることを避けます。 */
  fallbackColumnLimit?: number;
};

type Workbook = GC.Spread.Sheets.Workbook;
type Worksheet = GC.Spread.Sheets.Worksheet;

type UsedRangeSnapshot = {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
};

type SheetPosition = {
  sheetIndex: number;
  row: number;
  col: number;
};

type BoundColumnDefinition = {
  name?: unknown;
  displayName?: unknown;
};

const DEFAULT_MAX_RESULTS = 1000;
const DEFAULT_FALLBACK_ROW_LIMIT = 5000;
const DEFAULT_FALLBACK_COLUMN_LIMIT = 200;

type ScanRangeSnapshot = UsedRangeSnapshot & {
  source: 'usedRange' | 'usedRangeExpandedBySheetRange' | 'sheetRangeFallback';
  clippedRows: boolean;
  clippedColumns: boolean;
};

// SpreadJS の used range オブジェクトを strict TypeScript で安全に扱うための型ガードです。
function isUsedRangeSnapshot(value: unknown): value is UsedRangeSnapshot {
  if (!value || typeof value !== 'object') {
    return false;
  }

  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.row === 'number' &&
    typeof candidate.col === 'number' &&
    typeof candidate.rowCount === 'number' &&
    typeof candidate.colCount === 'number'
  );
}

function getUsedDataRange(sheet: Worksheet): ScanRangeSnapshot | null {
  const range: unknown = sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data);

  if (!isUsedRangeSnapshot(range) || range.rowCount <= 0 || range.colCount <= 0) {
    return null;
  }

  return {
    ...range,
    source: 'usedRange',
    clippedRows: false,
    clippedColumns: false,
  };
}

function getSheetName(sheet: Worksheet, fallbackIndex: number): string {
  const name = sheet.name();
  return name || `Sheet${fallbackIndex + 1}`;
}

function isBoundColumnDefinition(value: unknown): value is BoundColumnDefinition {
  return Boolean(value) && typeof value === 'object';
}

function getBoundColumns(sheet: Worksheet): BoundColumnDefinition[] {
  const sheetWithBinding = sheet as Worksheet & {
    bindColumns?: () => unknown;
  };

  if (typeof sheetWithBinding.bindColumns !== 'function') {
    return [];
  }

  const columns = sheetWithBinding.bindColumns();
  if (!Array.isArray(columns)) {
    return [];
  }

  return columns.map((column) => (isBoundColumnDefinition(column) ? column : {}));
}

function getBoundColumnName(boundColumns: BoundColumnDefinition[], columnIndex: number): string | null {
  const column = boundColumns[columnIndex];
  if (!column) {
    return null;
  }

  const name = typeof column.name === 'string' ? column.name : '';
  if (name) {
    return name;
  }

  const displayName = typeof column.displayName === 'string' ? column.displayName : '';
  return displayName || null;
}

function toResultAddress(row: number, col: number, boundColumns: BoundColumnDefinition[]): string {
  const boundColumnName = getBoundColumnName(boundColumns, col);
  if (boundColumnName) {
    return `${boundColumnName}[${row}]`;
  }

  return toA1Address(row, col);
}

function normalizeMaxResults(maxResults: number | undefined): number {
  if (maxResults === undefined || !Number.isFinite(maxResults)) {
    return DEFAULT_MAX_RESULTS;
  }

  return Math.max(1, Math.floor(maxResults));
}

function isAfterPosition(candidate: SheetPosition, origin: SheetPosition): boolean {
  if (candidate.sheetIndex !== origin.sheetIndex) {
    return candidate.sheetIndex > origin.sheetIndex;
  }

  if (candidate.row !== origin.row) {
    return candidate.row > origin.row;
  }

  return candidate.col > origin.col;
}

function getActivePosition(spread: Workbook): SheetPosition {
  const sheetIndex = Math.max(0, spread.getActiveSheetIndex());
  const sheet = spread.getSheet(sheetIndex);

  if (!sheet) {
    return { sheetIndex: 0, row: 0, col: -1 };
  }

  return {
    sheetIndex,
    row: sheet.getActiveRowIndex(),
    col: sheet.getActiveColumnIndex(),
  };
}

function normalizeRuntimeOptions(options: SearchRuntimeOptions | undefined): Required<SearchRuntimeOptions> {
  return {
    debug: options?.debug ?? false,
    fallbackToSheetRange: options?.fallbackToSheetRange ?? true,
    fallbackRowLimit: Math.max(1, Math.floor(options?.fallbackRowLimit ?? DEFAULT_FALLBACK_ROW_LIMIT)),
    fallbackColumnLimit: Math.max(1, Math.floor(options?.fallbackColumnLimit ?? DEFAULT_FALLBACK_COLUMN_LIMIT)),
  };
}

function getSheetScanRange(sheet: Worksheet, options: Required<SearchRuntimeOptions>): ScanRangeSnapshot | null {
  const usedRange = getUsedDataRange(sheet);
  if (usedRange) {
    if (!options.fallbackToSheetRange) {
      return usedRange;
    }

    const sheetRowCount = sheet.getRowCount();
    const sheetColumnCount = sheet.getColumnCount();
    const fallbackRowCount = Math.min(sheetRowCount, options.fallbackRowLimit);
    const fallbackColumnCount = Math.min(sheetColumnCount, options.fallbackColumnLimit);
    const row = Math.min(usedRange.row, 0);
    const col = Math.min(usedRange.col, 0);
    const rowEnd = Math.max(usedRange.row + usedRange.rowCount, fallbackRowCount);
    const colEnd = Math.max(usedRange.col + usedRange.colCount, fallbackColumnCount);

    return {
      row,
      col,
      rowCount: rowEnd - row,
      colCount: colEnd - col,
      source: 'usedRangeExpandedBySheetRange',
      clippedRows: sheetRowCount > fallbackRowCount,
      clippedColumns: sheetColumnCount > fallbackColumnCount,
    };
  }

  if (!options.fallbackToSheetRange) {
    return null;
  }

  const sheetRowCount = sheet.getRowCount();
  const sheetColumnCount = sheet.getColumnCount();
  const rowCount = Math.min(sheetRowCount, options.fallbackRowLimit);
  const colCount = Math.min(sheetColumnCount, options.fallbackColumnLimit);

  if (rowCount <= 0 || colCount <= 0) {
    return null;
  }

  return {
    row: 0,
    col: 0,
    rowCount,
    colCount,
    source: 'sheetRangeFallback',
    clippedRows: sheetRowCount > rowCount,
    clippedColumns: sheetColumnCount > colCount,
  };
}

function debugLog(options: Required<SearchRuntimeOptions>, message: string, details?: unknown): void {
  if (!options.debug || typeof console === 'undefined') {
    return;
  }

  if (details === undefined) {
    console.info(`[SpreadSearch] ${message}`);
    return;
  }

  console.info(`[SpreadSearch] ${message}`, details);
}

// getText の表示文字列を対象に、全シートをシート順・行方向順で走査します。
function* scanWorkbook(
  spread: Workbook,
  query: string,
  options: SearchOptions,
  runtimeOptions: Required<SearchRuntimeOptions>,
): Generator<SearchHit> {
  const sheetCount = spread.getSheetCount();

  debugLog(runtimeOptions, 'scan started', {
    query,
    options,
    sheetCount,
  });

  for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex += 1) {
    const sheet = spread.getSheet(sheetIndex);
    if (!sheet) {
      debugLog(runtimeOptions, 'sheet skipped: missing sheet instance', { sheetIndex });
      continue;
    }

    const sheetName = getSheetName(sheet, sheetIndex);
    const scanRange = getSheetScanRange(sheet, runtimeOptions);
    if (!scanRange) {
      debugLog(runtimeOptions, 'sheet skipped: no used range and no fallback range', {
        sheetIndex,
        sheetName,
        rowCount: sheet.getRowCount(),
        columnCount: sheet.getColumnCount(),
      });
      continue;
    }

    const rowEnd = scanRange.row + scanRange.rowCount;
    const colEnd = scanRange.col + scanRange.colCount;
    const boundColumns = getBoundColumns(sheet);
    let scannedCells = 0;
    let nonEmptyCells = 0;
    let hitCount = 0;

    debugLog(runtimeOptions, 'sheet scan range', {
      sheetIndex,
      sheetName,
      scanRange,
      boundColumnCount: boundColumns.length,
      sheetRowCount: sheet.getRowCount(),
      sheetColumnCount: sheet.getColumnCount(),
    });

    for (let row = scanRange.row; row < rowEnd; row += 1) {
      for (let col = scanRange.col; col < colEnd; col += 1) {
        scannedCells += 1;
        const text = String(sheet.getText(row, col) ?? '');
        if (text === '') {
          continue;
        }

        nonEmptyCells += 1;
        if (!matchesSearch(text, query, options)) {
          continue;
        }

        hitCount += 1;
        yield {
          sheetIndex,
          sheetName,
          row,
          col,
          address: toResultAddress(row, col, boundColumns),
          text,
        };
      }
    }

    debugLog(runtimeOptions, 'sheet scan finished', {
      sheetIndex,
      sheetName,
      scannedCells,
      nonEmptyCells,
      hitCount,
    });
  }
}

// 全件検索は結果一覧を作るための処理です。maxResults 到達時はそこで打ち切ります。
export function searchAllSheets(
  spread: Workbook,
  query: string,
  options: SearchOptions,
  maxResults?: number,
  runtimeOptions?: SearchRuntimeOptions,
): SearchAllResult {
  const normalizedRuntimeOptions = normalizeRuntimeOptions(runtimeOptions);

  if (query.length === 0) {
    debugLog(normalizedRuntimeOptions, 'search aborted: empty query');
    return {
      results: [],
      truncated: false,
      message: '検索文字列を入力してください。',
    };
  }

  const limit = normalizeMaxResults(maxResults);
  const results: SearchHit[] = [];

  for (const hit of scanWorkbook(spread, query, options, normalizedRuntimeOptions)) {
    results.push(hit);

    if (results.length >= limit) {
      debugLog(normalizedRuntimeOptions, 'search truncated', {
        limit,
        results,
      });
      return {
        results,
        truncated: true,
        message: `${limit}件に達したため打ち切りました。`,
      };
    }
  }

  debugLog(normalizedRuntimeOptions, 'search finished', {
    resultCount: results.length,
  });

  return {
    results,
    truncated: false,
    message: results.length === 0 ? '見つかりません。' : null,
  };
}

// 次を検索は結果一覧とは独立し、現在のアクティブセルの次から一周だけ探します。
export function findNextInWorkbook(
  spread: Workbook,
  query: string,
  options: SearchOptions,
  runtimeOptions?: SearchRuntimeOptions,
): FindNextResult {
  const normalizedRuntimeOptions = normalizeRuntimeOptions(runtimeOptions);

  if (query.length === 0) {
    debugLog(normalizedRuntimeOptions, 'find next aborted: empty query');
    return {
      hit: null,
      message: '検索文字列を入力してください。',
    };
  }

  const origin = getActivePosition(spread);
  let firstWrappedHit: SearchHit | null = null;

  debugLog(normalizedRuntimeOptions, 'find next started', {
    query,
    options,
    origin,
  });

  for (const hit of scanWorkbook(spread, query, options, normalizedRuntimeOptions)) {
    if (isAfterPosition(hit, origin)) {
      debugLog(normalizedRuntimeOptions, 'find next hit', hit);
      return { hit, message: null };
    }

    firstWrappedHit ??= hit;
  }

  if (firstWrappedHit) {
    debugLog(normalizedRuntimeOptions, 'find next wrapped hit', firstWrappedHit);
    return { hit: firstWrappedHit, message: null };
  }

  debugLog(normalizedRuntimeOptions, 'find next finished: no hit');
  return { hit: null, message: '見つかりません。' };
}

// 結果行クリックや次を検索で使う共通ジャンプ処理です。
export function jumpToHit(spread: Workbook, hit: SearchHit): JumpResult {
  const sheet = spread.getSheet(hit.sheetIndex);

  if (!sheet) {
    return { ok: false, message: 'シートが見つかりません。' };
  }

  if (
    hit.row < 0 ||
    hit.col < 0 ||
    hit.row >= sheet.getRowCount() ||
    hit.col >= sheet.getColumnCount()
  ) {
    return { ok: false, message: 'セルが見つかりません。' };
  }

  spread.setActiveSheetIndex(hit.sheetIndex);
  sheet.setActiveCell(hit.row, hit.col);
  sheet.showCell(
    hit.row,
    hit.col,
    GC.Spread.Sheets.VerticalPosition.center,
    GC.Spread.Sheets.HorizontalPosition.center,
  );

  return { ok: true, message: null };
}
