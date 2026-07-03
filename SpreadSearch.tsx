'use client';

/**
 * このファイルは画面に表示される検索 UI の本体です。
 * ここには固定トリガーボタン、モードレス検索ダイアログ、
 * ドラッグ移動、検索条件入力、結果一覧、結果クリック時のジャンプ処理を書いています。
 */

import type * as GC from '@grapecity/spread-sheets';
import type { CSSProperties, KeyboardEvent, PointerEvent } from 'react';
import { useEffect, useRef, useState } from 'react';
import {
  closeSpreadSearch,
  openSpreadSearch,
  setAllSearchResult,
  setSearchMessage,
  setSearchOption,
  setSearchQuery,
  setSelectedSearchIndex,
  useSpreadSearchStore,
} from './searchStore';
import { findNextInWorkbook, jumpToHit, searchAllSheets } from './spread-search';
import styles from './spreadSearch.module.css';

export type SpreadSearchProps = {
  /** 呼び出し時点の Workbook を返す getter。ホスト側で Workbook を useRef 等に保持して渡します。 */
  getSpread: () => GC.Spread.Sheets.Workbook | null | undefined;
  /** true の場合、画面右下に検索ダイアログを開く固定ボタンを表示します。 */
  showTrigger?: boolean;
  /** 固定トリガーボタンに表示する文字列です。 */
  triggerLabel?: string;
  /** 全件検索で返す最大件数です。超えた場合は打ち切りメッセージを出します。 */
  maxResults?: number;
  /** 既存画面より前面に出すための z-index です。 */
  zIndex?: number;
};

type DialogPosition = {
  x: number;
  y: number;
};

type DragState = {
  pointerId: number;
  startClientX: number;
  startClientY: number;
  originX: number;
  originY: number;
};

const DEFAULT_MAX_RESULTS = 1000;
const DEFAULT_Z_INDEX = 1000;

function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(' ');
}

/**
 * ホスト画面に 1 タグ追加するためのコンポーネントです。
 * 状態は searchStore.ts に置き、検索処理は spread-search.ts に委譲します。
 */
export function SpreadSearch({
  getSpread,
  showTrigger = true,
  triggerLabel = '検索',
  maxResults = DEFAULT_MAX_RESULTS,
  zIndex = DEFAULT_Z_INDEX,
}: SpreadSearchProps) {
  const state = useSpreadSearchStore();
  const dialogRef = useRef<HTMLDivElement | null>(null);
  const queryInputRef = useRef<HTMLInputElement | null>(null);
  const dragStateRef = useRef<DragState | null>(null);
  const [position, setPosition] = useState<DialogPosition>({ x: 96, y: 72 });

  // ダイアログを開いた直後に検索文字列へフォーカスします。
  useEffect(() => {
    if (!state.isOpen) {
      return;
    }

    queryInputRef.current?.focus();
  }, [state.isOpen]);

  function clampPosition(nextPosition: DialogPosition): DialogPosition {
    if (typeof window === 'undefined') {
      return nextPosition;
    }

    const dialog = dialogRef.current;
    const width = dialog?.offsetWidth ?? 420;
    const height = dialog?.offsetHeight ?? 360;
    const margin = 8;

    return {
      x: Math.min(Math.max(nextPosition.x, margin), Math.max(window.innerWidth - width - margin, margin)),
      y: Math.min(Math.max(nextPosition.y, margin), Math.max(window.innerHeight - height - margin, margin)),
    };
  }

  // タイトルバーの pointer イベントだけで、外部ライブラリなしのドラッグ移動を実装します。
  function handleTitlePointerDown(event: PointerEvent<HTMLDivElement>): void {
    event.currentTarget.setPointerCapture(event.pointerId);
    dragStateRef.current = {
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      originX: position.x,
      originY: position.y,
    };
  }

  function handleTitlePointerMove(event: PointerEvent<HTMLDivElement>): void {
    const dragState = dragStateRef.current;
    if (!dragState || dragState.pointerId !== event.pointerId) {
      return;
    }

    setPosition(
      clampPosition({
        x: dragState.originX + event.clientX - dragState.startClientX,
        y: dragState.originY + event.clientY - dragState.startClientY,
      }),
    );
  }

  function handleTitlePointerUp(event: PointerEvent<HTMLDivElement>): void {
    const dragState = dragStateRef.current;
    if (dragState?.pointerId === event.pointerId) {
      dragStateRef.current = null;
    }
  }

  function getReadySpread(): GC.Spread.Sheets.Workbook | null {
    const spread = getSpread();
    if (!spread) {
      setSearchMessage('シートが初期化されていません。');
      return null;
    }

    return spread;
  }

  // 全件検索は検索結果リストを更新します。
  function handleSearchAll(): void {
    const spread = getReadySpread();
    if (!spread) {
      return;
    }

    const result = searchAllSheets(spread, state.query, state.options, maxResults);
    setAllSearchResult({
      query: state.query,
      options: state.options,
      results: result.results,
      truncated: result.truncated,
      message: result.message,
    });
  }

  // 次を検索は検索結果リストを変更せず、現在のアクティブセルの次から直接ジャンプします。
  function handleFindNext(): void {
    const spread = getReadySpread();
    if (!spread) {
      return;
    }

    const result = findNextInWorkbook(spread, state.query, state.options);
    if (!result.hit) {
      setSearchMessage(result.message);
      return;
    }

    const jumpResult = jumpToHit(spread, result.hit);
    setSearchMessage(jumpResult.message);
  }

  // 結果行クリックでは、検索時点のスナップショット位置へ再検証なしでジャンプします。
  function handleResultClick(index: number): void {
    const spread = getReadySpread();
    if (!spread || !state.results) {
      return;
    }

    const hit = state.results[index];
    if (!hit) {
      return;
    }

    const result = jumpToHit(spread, hit);
    if (result.ok) {
      setSelectedSearchIndex(index);
    }
    setSearchMessage(result.message);
  }

  function handleQueryKeyDown(event: KeyboardEvent<HTMLInputElement>): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      handleFindNext();
      return;
    }

    if (event.key === 'Escape') {
      closeSpreadSearch();
    }
  }

  const dialogStyle: CSSProperties = {
    left: position.x,
    top: position.y,
    zIndex,
  };

  const triggerStyle: CSSProperties = {
    zIndex,
  };

  const resultCount = state.results?.length ?? 0;

  return (
    <>
      {showTrigger ? (
        <button
          className={styles['sjs-search-trigger']}
          onClick={openSpreadSearch}
          style={triggerStyle}
          type="button"
        >
          {triggerLabel}
        </button>
      ) : null}

      {state.isOpen ? (
        <div
          aria-label="SpreadJS 全シート検索"
          className={styles['sjs-search-dialog']}
          ref={dialogRef}
          role="dialog"
          style={dialogStyle}
        >
          <div
            className={styles['sjs-search-titlebar']}
            onPointerDown={handleTitlePointerDown}
            onPointerMove={handleTitlePointerMove}
            onPointerUp={handleTitlePointerUp}
          >
            <span>検索</span>
            <button
              aria-label="検索を閉じる"
              className={styles['sjs-search-close-button']}
              onClick={closeSpreadSearch}
              type="button"
            >
              ×
            </button>
          </div>

          <div className={styles['sjs-search-body']}>
            <label className={styles['sjs-search-field']}>
              <span>検索する文字列</span>
              <input
                className={styles['sjs-search-input']}
                onChange={(event) => setSearchQuery(event.target.value)}
                onKeyDown={handleQueryKeyDown}
                ref={queryInputRef}
                type="text"
                value={state.query}
              />
            </label>

            <div className={styles['sjs-search-options']}>
              <label className={styles['sjs-search-checkbox-row']}>
                <input
                  checked={state.options.exactMatch}
                  onChange={(event) => setSearchOption('exactMatch', event.target.checked)}
                  type="checkbox"
                />
                <span>完全一致</span>
              </label>
              <label className={styles['sjs-search-checkbox-row']}>
                <input
                  checked={state.options.matchByte}
                  onChange={(event) => setSearchOption('matchByte', event.target.checked)}
                  type="checkbox"
                />
                <span>半角と全角を区別する</span>
              </label>
              <label className={styles['sjs-search-checkbox-row']}>
                <input
                  checked={state.options.matchCase}
                  onChange={(event) => setSearchOption('matchCase', event.target.checked)}
                  type="checkbox"
                />
                <span>大文字と小文字を区別する</span>
              </label>
            </div>

            <div className={styles['sjs-search-actions']}>
              <button
                className={styles['sjs-search-primary-button']}
                onClick={handleSearchAll}
                type="button"
              >
                全件検索
              </button>
              <button
                className={styles['sjs-search-secondary-button']}
                onClick={handleFindNext}
                type="button"
              >
                次を検索
              </button>
              <button
                className={styles['sjs-search-secondary-button']}
                onClick={closeSpreadSearch}
                type="button"
              >
                キャンセル
              </button>
            </div>

            {state.message ? (
              <div className={styles['sjs-search-message']} role="status">
                {state.message}
              </div>
            ) : null}

            <section className={styles['sjs-search-results']}>
              <div className={styles['sjs-search-results-heading']}>
                <span>検索結果</span>
                <span>{state.results === null ? '未検索' : `${resultCount}件`}</span>
              </div>

              <div className={styles['sjs-search-table-wrap']}>
                <table className={styles['sjs-search-table']}>
                  <thead>
                    <tr>
                      <th scope="col">シート名</th>
                      <th scope="col">セル</th>
                      <th scope="col">値</th>
                    </tr>
                  </thead>
                  <tbody>
                    {state.results && state.results.length > 0 ? (
                      state.results.map((hit, index) => (
                        <tr
                          className={classNames(
                            styles['sjs-search-result-row'],
                            state.selectedIndex === index && styles['sjs-search-result-row-selected'],
                          )}
                          key={`${hit.sheetIndex}:${hit.row}:${hit.col}:${index}`}
                          onClick={() => handleResultClick(index)}
                        >
                          <td title={hit.sheetName}>{hit.sheetName}</td>
                          <td>{hit.address}</td>
                          <td title={hit.text}>{hit.text}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td className={styles['sjs-search-empty-cell']} colSpan={3}>
                          {state.results === null ? '検索結果はまだありません。' : '該当するセルはありません。'}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        </div>
      ) : null}
    </>
  );
}

export default SpreadSearch;
