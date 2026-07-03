'use client';

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
