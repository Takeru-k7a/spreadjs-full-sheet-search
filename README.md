# SpreadJS 全シート検索コンポーネント

React 18+ と `@grapecity/spread-sheets` だけに依存する自己完結フォルダです。フォルダごとコピーし、既存画面の JSX に `SpreadSearch` を 1 タグ追加すると、全シート検索、結果一覧、セルジャンプ、モードレスな検索ダイアログを使えます。

## ファイル構成

```text
spread-search/
  index.ts
  SpreadSearch.tsx
  searchStore.ts
  spread-search.ts
  textNormalize.ts
  spreadSearch.module.css
  README.md
  demo/SpreadSearchDemo.tsx
```

検索エンジンのファイルは `spread-search.ts` です。Windows などの大文字小文字を区別しない環境では `SpreadSearch.tsx` と `spreadSearch.ts` が TypeScript の解決順で衝突するため、実利用時に壊れない名前にしています。

## 前提

- React 18 以上
- `@grapecity/spread-sheets` v18 系
- CSS Modules を読み込める React/Next.js/Vite 等の環境
- Workbook はホスト側で作成済みで、`getSpread` から返す

コンポーネント内では `new GC.Spread.Sheets.Workbook(...)` を行いません。既存ストア、UI ライブラリ、Tailwind、Redux、Zustand、localStorage には依存しません。

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

`showTrigger` の既定値は `true` です。画面右下に固定の「検索」ボタンが表示され、クリックすると検索ダイアログが開きます。

## 自前ボタンから開く

```tsx
import { openSpreadSearch, SpreadSearch } from './components/spread-search';

function ExistingScreen() {
  return (
    <>
      <button type="button" onClick={openSpreadSearch}>
        検索
      </button>

      <SpreadSearch getSpread={() => spreadRef.current} showTrigger={false} />
    </>
  );
}
```

`closeSpreadSearch()` と `toggleSpreadSearch()` も export しています。

## Props

```ts
type SpreadSearchProps = {
  getSpread: () => GC.Spread.Sheets.Workbook | null | undefined;
  showTrigger?: boolean;   // default: true
  triggerLabel?: string;   // default: "検索"
  maxResults?: number;     // default: 1000
  zIndex?: number;         // default: 1000
};
```

`getSpread()` が `null` または `undefined` を返した場合は、エラーにせず「シートが初期化されていません。」と表示します。

## 検索仕様

- 対象は全シート固定です。
- `sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data)` の範囲を行方向に走査します。
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
