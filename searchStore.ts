'use client';

/**
 * このファイルは検索 UI の状態を保持する同梱ミニストアです。
 * React 以外の状態管理ライブラリを使わず、モジュールスコープの singleton と
 * useSyncExternalStore だけで、開閉状態、検索条件、結果一覧、選択行、メッセージを管理します。
 */

import { useSyncExternalStore } from 'react';

export type SearchOptions = {
  exactMatch: boolean;
  matchByte: boolean;
  matchCase: boolean;
};

export type SearchHit = {
  sheetIndex: number;
  sheetName: string;
  row: number;
  col: number;
  address: string;
  text: string;
};

export type SearchState = {
  isOpen: boolean;
  query: string;
  options: SearchOptions;
  results: SearchHit[] | null;
  truncated: boolean;
  searchedAt: number | null;
  selectedIndex: number | null;
  message: string | null;
};

type Listener = () => void;

// Excel 相当の初期値として、完全一致・半角全角区別・大文字小文字区別はすべて OFF です。
const defaultOptions: SearchOptions = {
  exactMatch: false,
  matchByte: false,
  matchCase: false,
};

let currentState: SearchState = {
  isOpen: false,
  query: '',
  options: defaultOptions,
  results: null,
  truncated: false,
  searchedAt: null,
  selectedIndex: null,
  message: null,
};

const listeners = new Set<Listener>();

// currentState を差し替えたあと、購読中の React コンポーネントへ更新通知します。
function emitChange(): void {
  listeners.forEach((listener) => listener());
}

function updateState(updater: (state: SearchState) => SearchState): void {
  currentState = updater(currentState);
  emitChange();
}

export function subscribeSearchStore(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

export function getSearchSnapshot(): SearchState {
  return currentState;
}

export function useSpreadSearchStore(): SearchState {
  return useSyncExternalStore(subscribeSearchStore, getSearchSnapshot, getSearchSnapshot);
}

// ここから下はホストアプリや SpreadSearch.tsx から呼ぶ状態更新 API です。
export function openSpreadSearch(): void {
  updateState((state) => ({ ...state, isOpen: true }));
}

export function closeSpreadSearch(): void {
  updateState((state) => ({ ...state, isOpen: false }));
}

export function toggleSpreadSearch(): void {
  updateState((state) => ({ ...state, isOpen: !state.isOpen }));
}

export function setSearchQuery(query: string): void {
  updateState((state) => ({ ...state, query }));
}

export function setSearchOption<Key extends keyof SearchOptions>(
  key: Key,
  value: SearchOptions[Key],
): void {
  updateState((state) => ({
    ...state,
    options: {
      ...state.options,
      [key]: value,
    },
  }));
}

export function setAllSearchResult(args: {
  query: string;
  options: SearchOptions;
  results: SearchHit[];
  truncated: boolean;
  message: string | null;
}): void {
  updateState((state) => ({
    ...state,
    query: args.query,
    options: { ...args.options },
    results: args.results,
    truncated: args.truncated,
    searchedAt: Date.now(),
    selectedIndex: null,
    message: args.message,
  }));
}

export function setSelectedSearchIndex(selectedIndex: number | null): void {
  updateState((state) => ({ ...state, selectedIndex }));
}

export function setSearchMessage(message: string | null): void {
  updateState((state) => ({ ...state, message }));
}
