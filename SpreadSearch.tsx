'use client';

/**
 * このファイルは画面に表示される検索 UI の本体です。
 * ここには固定トリガーボタン、モードレス検索ダイアログ、
 * ドラッグ移動、検索条件入力、結果一覧、結果クリック時のジャンプ処理を書いています。
 */

import type * as GC from '@grapecity/spread-sheets';
import {
  Box,
  Button,
  Checkbox,
  FormControlLabel,
  IconButton,
  Paper,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TextField,
  Typography,
} from '@mui/material';
import type { CSSProperties, KeyboardEvent, PointerEvent } from 'react';
import { useEffect, useRef } from 'react';
import {
  closeSpreadSearch,
  openSpreadSearch,
  setAllSearchResult,
  setSearchMessage,
  setSearchOption,
  setSearchQuery,
  setSelectedSearchIndex,
  setSpreadSearchPosition,
  type SearchDialogPosition,
  useSpreadSearchStore,
} from './searchStore';
import { findNextInWorkbook, jumpToHit, searchAllSheets, type SearchRuntimeOptions } from './spreadSearchEngine';

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
  /** 初期表示位置です。未指定時は左上寄りの既定位置に表示します。 */
  initialPosition?: SearchDialogPosition;
  /** true の場合、Ctrl+F / Cmd+F で検索ダイアログを開きます。 */
  enableShortcut?: boolean;
  /** true の場合、検索時のシート数・走査範囲・ヒット数を console に出します。 */
  debug?: boolean;
  /** usedRange が空のとき、sheet の行列数から fallback 走査するかどうかです。 */
  fallbackToSheetRange?: boolean;
  /** fallback 走査時に見る最大行数です。 */
  fallbackRowLimit?: number;
  /** fallback 走査時に見る最大列数です。 */
  fallbackColumnLimit?: number;
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

/**
 * ホスト画面に 1 タグ追加するためのコンポーネントです。
 * 状態は searchStore.ts に置き、検索処理は spreadSearchEngine.ts に委譲します。
 */
export function SpreadSearch({
  getSpread,
  showTrigger = true,
  triggerLabel = '検索',
  maxResults = DEFAULT_MAX_RESULTS,
  zIndex = DEFAULT_Z_INDEX,
  initialPosition,
  enableShortcut = true,
  debug = false,
  fallbackToSheetRange = true,
  fallbackRowLimit,
  fallbackColumnLimit,
}: SpreadSearchProps) {
  const state = useSpreadSearchStore();
  const dialogRef = useRef<HTMLDivElement | null>(null);
  const queryInputRef = useRef<HTMLInputElement | null>(null);
  const dragStateRef = useRef<DragState | null>(null);
  const dragCleanupRef = useRef<(() => void) | null>(null);
  const dragFrameRef = useRef<number | null>(null);
  const pendingDragPositionRef = useRef<SearchDialogPosition | null>(null);
  const didApplyInitialPositionRef = useRef(false);
  const runtimeOptions: SearchRuntimeOptions = {
    debug,
    fallbackToSheetRange,
    fallbackRowLimit,
    fallbackColumnLimit,
  };

  // ホスト側が初期位置を指定した場合、一度だけストアへ反映します。
  useEffect(() => {
    if (!initialPosition || didApplyInitialPositionRef.current) {
      return;
    }

    didApplyInitialPositionRef.current = true;
    setSpreadSearchPosition(clampPosition(initialPosition));
  }, [initialPosition]);

  // Ctrl+F / Cmd+F でブラウザ標準検索の代わりに SpreadJS 検索を開きます。
  useEffect(() => {
    if (!enableShortcut || typeof window === 'undefined') {
      return undefined;
    }

    function handleGlobalKeyDown(event: globalThis.KeyboardEvent): void {
      const isFindShortcut =
        (event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 'f';

      if (isFindShortcut) {
        event.preventDefault();
        event.stopPropagation();
        openSpreadSearch();
        return;
      }

      if (state.isOpen && event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        closeSpreadSearch();
      }
    }

    window.addEventListener('keydown', handleGlobalKeyDown, true);
    return () => {
      window.removeEventListener('keydown', handleGlobalKeyDown, true);
    };
  }, [enableShortcut, state.isOpen]);

  // ダイアログを開いた直後に検索文字列へフォーカスします。
  useEffect(() => {
    if (!state.isOpen) {
      return;
    }

    queryInputRef.current?.focus();
  }, [state.isOpen]);

  useEffect(() => {
    return () => {
      clearDragListeners();
    };
  }, []);

  function clampPosition(nextPosition: SearchDialogPosition): SearchDialogPosition {
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

  function clearDragListeners(flushPendingPosition = false): void {
    const pendingPosition = pendingDragPositionRef.current;
    dragCleanupRef.current?.();
    dragCleanupRef.current = null;
    if (dragFrameRef.current !== null && typeof window !== 'undefined') {
      window.cancelAnimationFrame(dragFrameRef.current);
      dragFrameRef.current = null;
    }
    pendingDragPositionRef.current = null;
    if (flushPendingPosition && pendingPosition) {
      setSpreadSearchPosition(pendingPosition);
    }
  }

  function scheduleDragPosition(nextPosition: SearchDialogPosition): void {
    pendingDragPositionRef.current = nextPosition;
    if (dragFrameRef.current !== null) {
      return;
    }

    dragFrameRef.current = window.requestAnimationFrame(() => {
      dragFrameRef.current = null;
      const pendingPosition = pendingDragPositionRef.current;
      pendingDragPositionRef.current = null;
      if (pendingPosition) {
        setSpreadSearchPosition(pendingPosition);
      }
    });
  }

  // pointer capture に依存せず、window 側で移動を追跡してドラッグ移動を安定させます。
  function handleTitlePointerDown(event: PointerEvent<HTMLDivElement>): void {
    if (typeof window === 'undefined') {
      return;
    }

    const target = event.target instanceof HTMLElement ? event.target : null;
    if (target?.closest('button')) {
      return;
    }

    event.preventDefault();
    clearDragListeners();
    dragStateRef.current = {
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      originX: state.position.x,
      originY: state.position.y,
    };

    function handleWindowPointerMove(nativeEvent: globalThis.PointerEvent): void {
      const dragState = dragStateRef.current;
      if (!dragState || dragState.pointerId !== nativeEvent.pointerId) {
        return;
      }

      nativeEvent.preventDefault();
      scheduleDragPosition(
        clampPosition({
          x: dragState.originX + nativeEvent.clientX - dragState.startClientX,
          y: dragState.originY + nativeEvent.clientY - dragState.startClientY,
        }),
      );
    }

    function handleWindowPointerEnd(nativeEvent: globalThis.PointerEvent): void {
      const dragState = dragStateRef.current;
      if (dragState && dragState.pointerId !== nativeEvent.pointerId) {
        return;
      }

      dragStateRef.current = null;
      clearDragListeners(true);
    }

    window.addEventListener('pointermove', handleWindowPointerMove, true);
    window.addEventListener('pointerup', handleWindowPointerEnd, true);
    window.addEventListener('pointercancel', handleWindowPointerEnd, true);
    dragCleanupRef.current = () => {
      window.removeEventListener('pointermove', handleWindowPointerMove, true);
      window.removeEventListener('pointerup', handleWindowPointerEnd, true);
      window.removeEventListener('pointercancel', handleWindowPointerEnd, true);
    };
  }

  function getReadySpread(): GC.Spread.Sheets.Workbook | null {
    const spread = getSpread();
    if (!spread) {
      if (debug && typeof console !== 'undefined') {
        console.info('[SpreadSearch] getSpread returned null or undefined');
      }
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

    const result = searchAllSheets(spread, state.query, state.options, maxResults, runtimeOptions);
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

    const result = findNextInWorkbook(spread, state.query, state.options, runtimeOptions);
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
    left: state.position.x,
    top: state.position.y,
    zIndex,
  };

  const triggerStyle: CSSProperties = {
    zIndex,
  };

  const resultCount = state.results?.length ?? 0;

  return (
    <>
      {showTrigger ? (
        <Button
          size="small"
          variant="outlined"
          onClick={() => openSpreadSearch()}
          sx={{
            position: 'fixed',
            right: 16,
            bottom: 16,
            zIndex: triggerStyle.zIndex,
            minWidth: 72,
            bgcolor: 'background.paper',
            boxShadow: 3,
          }}
          type="button"
        >
          {triggerLabel}
        </Button>
      ) : null}

      {state.isOpen ? (
        <Paper
          aria-label="SpreadJS 全シート検索"
          elevation={10}
          ref={dialogRef}
          role="dialog"
          sx={{
            position: 'fixed',
            display: 'grid',
            gridTemplateRows: 'auto minmax(0, 1fr)',
            left: dialogStyle.left,
            top: dialogStyle.top,
            zIndex: dialogStyle.zIndex,
            width: 'min(430px, calc(100vw - 16px))',
            maxHeight: 'calc(100vh - 16px)',
            overflow: 'hidden',
            border: 1,
            borderColor: 'divider',
            borderRadius: 1,
          }}
        >
          <Box
            onPointerDown={handleTitlePointerDown}
            sx={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              minHeight: 40,
              px: 1.5,
              py: 0.5,
              borderBottom: 1,
              borderColor: 'divider',
              bgcolor: 'grey.100',
              cursor: 'grab',
              userSelect: 'none',
              touchAction: 'none',
              '&:active': {
                cursor: 'grabbing',
              },
            }}
          >
            <Typography component="h2" variant="subtitle1" sx={{ fontWeight: 700 }}>
              検索
            </Typography>
            <IconButton
              aria-label="検索を閉じる"
              onClick={closeSpreadSearch}
              size="small"
              type="button"
            >
              ×
            </IconButton>
          </Box>

          <Stack spacing={1.25} sx={{ minHeight: 0, p: 1.5, overflow: 'auto' }}>
            <TextField
              inputRef={queryInputRef}
              label="検索する文字列"
              onChange={(event) => setSearchQuery(event.target.value)}
              onKeyDown={handleQueryKeyDown}
              size="small"
              type="text"
              value={state.query}
              fullWidth
            />

            <Stack spacing={0} sx={{ mt: 0.75 }}>
              <FormControlLabel
                control={
                  <Checkbox
                    checked={state.options.exactMatch}
                    onChange={(event) => setSearchOption('exactMatch', event.target.checked)}
                    size="small"
                    sx={{ p: 0.25, '& .MuiSvgIcon-root': { fontSize: 17 } }}
                  />
                }
                label={<Typography variant="body2">完全一致</Typography>}
                sx={{ m: 0, minHeight: 26 }}
              />
              <FormControlLabel
                control={
                  <Checkbox
                    checked={state.options.matchByte}
                    onChange={(event) => setSearchOption('matchByte', event.target.checked)}
                    size="small"
                    sx={{ p: 0.25, '& .MuiSvgIcon-root': { fontSize: 17 } }}
                  />
                }
                label={<Typography variant="body2">半角と全角を区別する</Typography>}
                sx={{ m: 0, minHeight: 26 }}
              />
              <FormControlLabel
                control={
                  <Checkbox
                    checked={state.options.matchCase}
                    onChange={(event) => setSearchOption('matchCase', event.target.checked)}
                    size="small"
                    sx={{ p: 0.25, '& .MuiSvgIcon-root': { fontSize: 17 } }}
                  />
                }
                label={<Typography variant="body2">大文字と小文字を区別する</Typography>}
                sx={{ m: 0, minHeight: 26 }}
              />
            </Stack>

            <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1 }}>
              <Button onClick={handleSearchAll} type="button" variant="contained">
                全件検索
              </Button>
              <Button onClick={handleFindNext} type="button" variant="outlined">
                次を検索
              </Button>
              <Button onClick={closeSpreadSearch} type="button" variant="outlined">
                閉じる
              </Button>
            </Box>

            {state.message ? (
              <Paper
                role="status"
                variant="outlined"
                sx={{
                  px: 1,
                  py: 0.75,
                  bgcolor: 'warning.50',
                  borderColor: 'warning.light',
                  color: 'warning.dark',
                }}
              >
                {state.message}
              </Paper>
            ) : null}

            <Box component="section" sx={{ display: 'grid', gap: 0.75, minHeight: 0, pt: 0.5 }}>
              <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1 }}>
                <Typography variant="subtitle2" sx={{ fontWeight: 700 }}>
                  検索結果
                </Typography>
                <Typography variant="caption" color="text.secondary">
                  {state.results === null ? '未検索' : `${resultCount}件`}
                </Typography>
              </Box>

              <TableContainer component={Paper} variant="outlined" sx={{ maxHeight: 230 }}>
                <Table stickyHeader size="small" sx={{ tableLayout: 'fixed' }}>
                  <TableHead>
                    <TableRow>
                      <TableCell sx={{ width: '34%', fontWeight: 700 }}>シート名</TableCell>
                      <TableCell sx={{ width: 72, fontWeight: 700 }}>セル</TableCell>
                      <TableCell sx={{ fontWeight: 700 }}>値</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {state.results && state.results.length > 0 ? (
                      state.results.map((hit, index) => (
                        <TableRow
                          hover
                          key={`${hit.sheetIndex}:${hit.row}:${hit.col}:${index}`}
                          onClick={() => handleResultClick(index)}
                          selected={state.selectedIndex === index}
                          sx={{ cursor: 'pointer' }}
                        >
                          <TableCell
                            title={hit.sheetName}
                            sx={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                          >
                            {hit.sheetName}
                          </TableCell>
                          <TableCell>{hit.address}</TableCell>
                          <TableCell
                            title={hit.text}
                            sx={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                          >
                            {hit.text}
                          </TableCell>
                        </TableRow>
                      ))
                    ) : (
                      <TableRow>
                        <TableCell colSpan={3} align="center" sx={{ color: 'text.secondary' }}>
                          {state.results === null ? '検索結果はまだありません。' : '該当するセルはありません。'}
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </TableContainer>
            </Box>
          </Stack>
        </Paper>
      ) : null}
    </>
  );
}

export default SpreadSearch;
