'use client';

import type { SearchOptions } from './searchStore';

const halfWidthKanaMap: Record<string, string> = {
  '｡': '。',
  '｢': '「',
  '｣': '」',
  '､': '、',
  '･': '・',
  'ｦ': 'ヲ',
  'ｧ': 'ァ',
  'ｨ': 'ィ',
  'ｩ': 'ゥ',
  'ｪ': 'ェ',
  'ｫ': 'ォ',
  'ｬ': 'ャ',
  'ｭ': 'ュ',
  'ｮ': 'ョ',
  'ｯ': 'ッ',
  'ｰ': 'ー',
  'ｱ': 'ア',
  'ｲ': 'イ',
  'ｳ': 'ウ',
  'ｴ': 'エ',
  'ｵ': 'オ',
  'ｶ': 'カ',
  'ｷ': 'キ',
  'ｸ': 'ク',
  'ｹ': 'ケ',
  'ｺ': 'コ',
  'ｻ': 'サ',
  'ｼ': 'シ',
  'ｽ': 'ス',
  'ｾ': 'セ',
  'ｿ': 'ソ',
  'ﾀ': 'タ',
  'ﾁ': 'チ',
  'ﾂ': 'ツ',
  'ﾃ': 'テ',
  'ﾄ': 'ト',
  'ﾅ': 'ナ',
  'ﾆ': 'ニ',
  'ﾇ': 'ヌ',
  'ﾈ': 'ネ',
  'ﾉ': 'ノ',
  'ﾊ': 'ハ',
  'ﾋ': 'ヒ',
  'ﾌ': 'フ',
  'ﾍ': 'ヘ',
  'ﾎ': 'ホ',
  'ﾏ': 'マ',
  'ﾐ': 'ミ',
  'ﾑ': 'ム',
  'ﾒ': 'メ',
  'ﾓ': 'モ',
  'ﾔ': 'ヤ',
  'ﾕ': 'ユ',
  'ﾖ': 'ヨ',
  'ﾗ': 'ラ',
  'ﾘ': 'リ',
  'ﾙ': 'ル',
  'ﾚ': 'レ',
  'ﾛ': 'ロ',
  'ﾜ': 'ワ',
  'ﾝ': 'ン',
};

const dakutenMap: Record<string, string> = {
  ウ: 'ヴ',
  カ: 'ガ',
  キ: 'ギ',
  ク: 'グ',
  ケ: 'ゲ',
  コ: 'ゴ',
  サ: 'ザ',
  シ: 'ジ',
  ス: 'ズ',
  セ: 'ゼ',
  ソ: 'ゾ',
  タ: 'ダ',
  チ: 'ヂ',
  ツ: 'ヅ',
  テ: 'デ',
  ト: 'ド',
  ハ: 'バ',
  ヒ: 'ビ',
  フ: 'ブ',
  ヘ: 'ベ',
  ホ: 'ボ',
  ワ: 'ヷ',
  ヰ: 'ヸ',
  ヱ: 'ヹ',
  ヲ: 'ヺ',
};

const handakutenMap: Record<string, string> = {
  ハ: 'パ',
  ヒ: 'ピ',
  フ: 'プ',
  ヘ: 'ペ',
  ホ: 'ポ',
};

export function toWidthInsensitive(input: string): string {
  const output: string[] = [];

  for (const char of input) {
    if (char === 'ﾞ') {
      const previous = output[output.length - 1];
      if (previous && dakutenMap[previous]) {
        output[output.length - 1] = dakutenMap[previous];
      } else {
        output.push('゛');
      }
      continue;
    }

    if (char === 'ﾟ') {
      const previous = output[output.length - 1];
      if (previous && handakutenMap[previous]) {
        output[output.length - 1] = handakutenMap[previous];
      } else {
        output.push('゜');
      }
      continue;
    }

    const codePoint = char.codePointAt(0);
    if (codePoint === undefined) {
      continue;
    }

    if (codePoint === 0x3000) {
      output.push(' ');
      continue;
    }

    if (codePoint >= 0xff01 && codePoint <= 0xff5e) {
      output.push(String.fromCodePoint(codePoint - 0xfee0));
      continue;
    }

    output.push(halfWidthKanaMap[char] ?? char);
  }

  return output.join('');
}

export function normalizeForSearch(input: string, options: SearchOptions): string {
  let text = input;

  if (!options.matchByte) {
    text = toWidthInsensitive(text);
  }

  if (!options.matchCase) {
    text = text.toLowerCase();
  }

  return text;
}

export function matchesSearch(text: string, query: string, options: SearchOptions): boolean {
  const normalizedText = normalizeForSearch(text, options);
  const normalizedQuery = normalizeForSearch(query, options);

  if (options.exactMatch) {
    return normalizedText === normalizedQuery;
  }

  return normalizedText.includes(normalizedQuery);
}

export function columnIndexToName(columnIndex: number): string {
  if (!Number.isInteger(columnIndex) || columnIndex < 0) {
    return '';
  }

  let current = columnIndex;
  let columnName = '';

  do {
    const remainder = current % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return columnName;
}

export function toA1Address(row: number, column: number): string {
  const columnName = columnIndexToName(column);
  return columnName ? `${columnName}${row + 1}` : '';
}
