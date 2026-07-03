'use client';

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

const DEFAULT_MAX_RESULTS = 1000;

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

function getUsedDataRange(sheet: Worksheet): UsedRangeSnapshot | null {
  const range: unknown = sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data);

  if (!isUsedRangeSnapshot(range) || range.rowCount <= 0 || range.colCount <= 0) {
    return null;
  }

  return range;
}

function getSheetName(sheet: Worksheet, fallbackIndex: number): string {
  const name = sheet.name();
  return name || `Sheet${fallbackIndex + 1}`;
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

function* scanWorkbook(
  spread: Workbook,
  query: string,
  options: SearchOptions,
): Generator<SearchHit> {
  const sheetCount = spread.getSheetCount();

  for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex += 1) {
    const sheet = spread.getSheet(sheetIndex);
    if (!sheet) {
      continue;
    }

    const usedRange = getUsedDataRange(sheet);
    if (!usedRange) {
      continue;
    }

    const rowEnd = usedRange.row + usedRange.rowCount;
    const colEnd = usedRange.col + usedRange.colCount;
    const sheetName = getSheetName(sheet, sheetIndex);

    for (let row = usedRange.row; row < rowEnd; row += 1) {
      for (let col = usedRange.col; col < colEnd; col += 1) {
        const text = String(sheet.getText(row, col) ?? '');
        if (text === '' || !matchesSearch(text, query, options)) {
          continue;
        }

        yield {
          sheetIndex,
          sheetName,
          row,
          col,
          address: toA1Address(row, col),
          text,
        };
      }
    }
  }
}

export function searchAllSheets(
  spread: Workbook,
  query: string,
  options: SearchOptions,
  maxResults?: number,
): SearchAllResult {
  if (query.length === 0) {
    return {
      results: [],
      truncated: false,
      message: '検索文字列を入力してください。',
    };
  }

  const limit = normalizeMaxResults(maxResults);
  const results: SearchHit[] = [];

  for (const hit of scanWorkbook(spread, query, options)) {
    results.push(hit);

    if (results.length >= limit) {
      return {
        results,
        truncated: true,
        message: `${limit}件に達したため打ち切りました。`,
      };
    }
  }

  return {
    results,
    truncated: false,
    message: results.length === 0 ? '見つかりません。' : null,
  };
}

export function findNextInWorkbook(
  spread: Workbook,
  query: string,
  options: SearchOptions,
): FindNextResult {
  if (query.length === 0) {
    return {
      hit: null,
      message: '検索文字列を入力してください。',
    };
  }

  const origin = getActivePosition(spread);
  let firstWrappedHit: SearchHit | null = null;

  for (const hit of scanWorkbook(spread, query, options)) {
    if (isAfterPosition(hit, origin)) {
      return { hit, message: null };
    }

    firstWrappedHit ??= hit;
  }

  return firstWrappedHit
    ? { hit: firstWrappedHit, message: null }
    : { hit: null, message: '見つかりません。' };
}

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
