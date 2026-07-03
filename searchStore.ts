'use client';

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
