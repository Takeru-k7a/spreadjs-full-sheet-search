# SpreadJS 全シート検索コンポーネント

React 18+ と `@grapecity/spread-sheets` だけに依存する自己完結フォルダです。フォルダごとコピーし、既存画面の JSX に `SpreadSearch` を 1 タグ追加すると、全シート検索、結果一覧、セルジャンプ、モードレスな検索ダイアログを使えます。

## ファイル構成

```text
spread-search/
  package.json
  index.ts
  SpreadSearch.tsx
  searchStore.ts
  spreadSearchEngine.ts
  textNormalize.ts
  README.md
  demo/SpreadSearchDemo.tsx
```

公開入口は `index.ts` です。`spreadSearchEngine.ts` は Workbook を走査する内部ロジックで、`SpreadSearch` コンポーネント自体は `index.ts` から export しています。

## 前提

- React 18 以上
- `@grapecity/spread-sheets` v18 系
- `@mui/material`
- `@emotion/react`
- `@emotion/styled`
- Workbook はホスト側で作成済みで、`getSpread` から返す

コンポーネント内では `new GC.Spread.Sheets.Workbook(...)` を行いません。既存ストア、Tailwind、Redux、Zustand、localStorage には依存しません。UI は MUI Material を使います。

未導入の場合はホストアプリ側で追加してください。

```bash
npm install @mui/material @emotion/react @emotion/styled
```

## 置き方

任意の場所へ `spread-search/` をコピーします。例:

```text
src/components/spread-search/
```

その場合の import 例:

```tsx
import { SpreadSearch } from './components/spread-search';
```

## 1 タグ追加で使う

```tsx
import type * as GC from '@grapecity/spread-sheets';
import { useRef } from 'react';
import { SpreadSearch } from './components/spread-search';

function ExistingScreen() {
  const spreadRef = useRef<GC.Spread.Sheets.Workbook | null>(null);

  return (
    <>
      {/* 既存の SpreadJS ホスト */}
      <div id="spread-host" />

      {/* 追加するのはこの 1 タグ */}
      <SpreadSearch getSpread={() => spreadRef.current} />
    </>
  );
}
```

`showTrigger` の既定値は `true` です。画面右下に固定の「検索」ボタンが表示され、クリックすると検索ダイアログが開きます。`enableShortcut` の既定値も `true` なので、Ctrl+F / Cmd+F でも開けます。

## Ctrl+F だけで使う

画面上の検索ボタンを出したくない場合は `showTrigger={false}` を指定します。この場合でも、既定で Ctrl+F / Cmd+F から検索ダイアログを開けます。

```tsx
<SpreadSearch getSpread={() => spreadRef.current} showTrigger={false} />
```

ブラウザ標準検索を優先したい画面では `enableShortcut={false}` を指定してください。

## 自前ボタンから開く

```tsx
import { closeSpreadSearch, openSpreadSearch, SpreadSearch } from './components/spread-search';

function ExistingScreen() {
  return (
    <>
      <button type="button" onClick={() => openSpreadSearch()}>
        検索
      </button>
      <button type="button" onClick={() => closeSpreadSearch()}>
        閉じる
      </button>

      <SpreadSearch getSpread={() => spreadRef.current} showTrigger={false} />
    </>
  );
}
```

`closeSpreadSearch()`、`toggleSpreadSearch()`、`setSpreadSearchPosition()`、`resetSpreadSearchPosition()` も export しています。

## Props

```ts
type SpreadSearchProps = {
  getSpread: () => GC.Spread.Sheets.Workbook | null | undefined;
  showTrigger?: boolean;   // default: true
  triggerLabel?: string;   // default: "検索"
  maxResults?: number;     // default: 1000
  zIndex?: number;         // default: 1000
  initialPosition?: { x: number; y: number };
  enableShortcut?: boolean; // default: true
  debug?: boolean;         // default: false
  fallbackToSheetRange?: boolean; // default: true
  fallbackRowLimit?: number;      // default: 5000
  fallbackColumnLimit?: number;   // default: 200
};
```

`getSpread()` が `null` または `undefined` を返した場合は、エラーにせず「シートが初期化されていません。」と表示します。

検索ダイアログはタイトルバーをドラッグして移動できます。移動後の位置はストアに残るため、閉じて再度開いても同じ位置に表示されます。初期位置を指定したい場合は次のようにします。

```tsx
<SpreadSearch getSpread={() => spreadRef.current} initialPosition={{ x: 320, y: 80 }} />
```

ホスト側のボタンやショートカットから位置を変える場合は `setSpreadSearchPosition({ x, y })`、既定位置に戻す場合は `resetSpreadSearchPosition()` を使います。閉じる場合はダイアログ右上の `×`、下部の「閉じる」、Esc キー、または `closeSpreadSearch()` を使えます。

`debug` を `true` にすると、検索実行時に `console.info` へ Workbook のシート数、各シートの走査範囲、非空セル数、ヒット数を出します。検索結果が出ない画面では、まず次のように一時的に有効化してください。

```tsx
<SpreadSearch getSpread={() => spreadRef.current} debug />
```

`getUsedRange(GC.Spread.Sheets.UsedRangeType.data)` が空、または実データ範囲より狭いホスト画面でも検索できるよう、既定で `fallbackToSheetRange` が有効です。この場合は `sheet.getRowCount()` / `sheet.getColumnCount()` を使って先頭側の走査範囲も含めますが、巨大シートで固まらないよう `fallbackRowLimit` と `fallbackColumnLimit` で上限をかけます。

## 検索仕様

- 対象は全シート固定です。
- まず `sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data)` の範囲を行方向に走査します。
- 既定では usedRange に加えて、シート行列数の先頭範囲も fallback 走査します。
- 検索対象は `sheet.getText(row, col)` の表示文字列です。
- 非表示の行、列、シートも検索対象に含みます。
- 全件検索の結果は、検索時点のスナップショットとして保持します。
- 結果行クリック時は、再検索せず記録済みのシート、行、列へジャンプします。
- 「次を検索」は、全件検索の結果リストとは独立して、現在の条件で都度検索します。

半角全角を区別しない場合、全角英数記号、全角スペース、半角カタカナだけを自前変換します。`String.prototype.normalize('NFKC')` は使っていないため、`①` と `1` は同一視しません。

## スコープ外

- 置換
- ワイルドカード、正規表現検索
- 数式、コメント、タグ内検索
- ヒットセルの背景色変更などのハイライト
- 検索範囲の切り替え
- localStorage 等への永続化

## デモ

`demo/SpreadSearchDemo.tsx` は、複数シート、全角半角、大文字小文字、完全一致、`maxResults` 打ち切りを確認するための最小ページです。

```tsx
import { SpreadSearchDemo } from './components/spread-search/demo/SpreadSearchDemo';

export default function Page() {
  return <SpreadSearchDemo />;
}
```

SpreadJS 本体の CSS は、ホストアプリ側で通常どおり読み込んでください。
