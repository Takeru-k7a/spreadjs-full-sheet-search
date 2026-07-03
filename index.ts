'use client';

/**
 * このファイルは公開 API の入口です。
 * ホストアプリは基本的にここから SpreadSearch コンポーネント、
 * open/close/toggle 関数、検索関連の型とユーティリティを import します。
 */

export { SpreadSearch, type SpreadSearchProps } from './SpreadSearch';
export {
  closeSpreadSearch,
  openSpreadSearch,
  toggleSpreadSearch,
  useSpreadSearchStore,
} from './searchStore';
export type { SearchHit, SearchOptions, SearchState } from './searchStore';
export {
  findNextInWorkbook,
  jumpToHit,
  searchAllSheets,
} from './spread-search';
export type { FindNextResult, JumpResult, SearchAllResult } from './spread-search';
export {
  columnIndexToName,
  matchesSearch,
  normalizeForSearch,
  toA1Address,
  toWidthInsensitive,
} from './textNormalize';
