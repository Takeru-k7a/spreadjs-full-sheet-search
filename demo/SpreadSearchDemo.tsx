'use client';

/**
 * このファイルは動作確認用の最小デモページです。
 * 複数シート、半角全角、大文字小文字、完全一致、maxResults 打ち切りを
 * 手元で確認できるサンプル Workbook を作ります。
 */

import * as GC from '@grapecity/spread-sheets';
import { useCallback, useRef } from 'react';
import { openSpreadSearch, SpreadSearch } from '../index';

type Workbook = GC.Spread.Sheets.Workbook;
type Worksheet = GC.Spread.Sheets.Worksheet;

// デモ用シートへ二次元配列を流し込み、最低限の行列数と列幅を整えます。
function setSheetData(sheet: Worksheet, name: string, values: unknown[][]): void {
  sheet.name(name);
  sheet.setRowCount(Math.max(values.length + 4, 24));
  sheet.setColumnCount(8);
  sheet.setArray(0, 0, values);
  sheet.setColumnWidth(0, 150);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 180);
}

// 受け入れ条件を確認しやすいよう、検索パターン別のシートを作ります。
function fillDemoWorkbook(workbook: Workbook): void {
  workbook.suspendPaint();

  const customers = workbook.getSheet(0);
  const widthCases = workbook.getSheet(1);
  const limitCases = workbook.getSheet(2);

  if (customers) {
    setSheetData(customers, '顧客', [
      ['会社名', '地域', 'コード', '備考'],
      ['ABC商事', '東京都', 'A1', '大文字小文字の確認用'],
      ['東京食品', '大阪府', 'B2', '完全一致 OFF では「東京」にヒット'],
      ['関西販売', '京都府', 'Ｃ３', '全角英数の確認用'],
    ]);
  }

  if (widthCases) {
    setSheetData(widthCases, '幅変換', [
      ['項目', '全角', '半角', '備考'],
      ['国名', 'アメリカ', 'ｱﾒﾘｶ', '半角全角区別 OFF で一致'],
      ['コード', 'Ａ１', 'A1', '半角全角区別 OFF で一致'],
      ['濁点', 'ガ', 'ｶﾞ', '濁点合成の確認用'],
      ['半濁点', 'パ', 'ﾊﾟ', '半濁点合成の確認用'],
      ['丸数字', '①', '1', 'NFKC を使わない確認用'],
    ]);
  }

  if (limitCases) {
    limitCases.name('上限確認');
    limitCases.setRowCount(1050);
    limitCases.setColumnCount(4);
    limitCases.setArray(0, 0, [['No', '値', '分類', '備考']]);
    limitCases.setColumnWidth(0, 80);
    limitCases.setColumnWidth(1, 180);
    limitCases.setColumnWidth(2, 120);
    limitCases.setColumnWidth(3, 220);

    for (let row = 1; row <= 1020; row += 1) {
      limitCases.setValue(row, 0, row);
      limitCases.setValue(row, 1, `maxResults-${row}`);
      limitCases.setValue(row, 2, row % 2 === 0 ? 'even' : 'odd');
      limitCases.setValue(row, 3, 'maxResults の打ち切り確認用データ');
    }
  }

  workbook.resumePaint();
}

// README の例として使える、Workbook getter と SpreadSearch の最小構成です。
export function SpreadSearchDemo() {
  const spreadRef = useRef<Workbook | null>(null);

  const setHostRef = useCallback((node: HTMLDivElement | null) => {
    if (!node || spreadRef.current) {
      return;
    }

    const workbook = new GC.Spread.Sheets.Workbook(node, {
      sheetCount: 3,
      tabStripVisible: true,
      newTabVisible: false,
    });

    spreadRef.current = workbook;
    fillDemoWorkbook(workbook);
  }, []);

  return (
    <main
      style={{
        display: 'grid',
        gridTemplateRows: 'auto minmax(0, 1fr)',
        gap: 12,
        height: '100vh',
        minHeight: 480,
        padding: 12,
      }}
    >
      <header style={{ display: 'flex', flexWrap: 'wrap', gap: 8, justifyContent: 'space-between' }}>
        <h1 style={{ fontSize: 18, margin: 0 }}>SpreadSearch Demo</h1>
        <button onClick={openSpreadSearch} type="button">
          自前ボタンで検索を開く
        </button>
      </header>

      <div ref={setHostRef} style={{ minHeight: 0, border: '1px solid #c8d3df' }} />

      <SpreadSearch getSpread={() => spreadRef.current} />
    </main>
  );
}

export default SpreadSearchDemo;
