import type { GridBlock, StyleValue } from './types';

export function textBlock(
  text: string,
  options?: {
    style?: StyleValue;
    colSpan?: number | 'remaining';
    height?: number;
  },
): GridBlock {
  return {
    type: 'grid',
    rows: [
      {
        height: options?.height,
        cells: [
          {
            value: text,
            style: options?.style,
            colSpan: options?.colSpan,
          },
        ],
      },
    ],
  };
}

export function spacerBlock(rows: number = 1): GridBlock {
  return {
    type: 'grid',
    rows: Array.from({ length: rows }).map(() => ({ cells: [] })),
  };
}

export function dividerBlock(
  rows: number = 1,
  options?: {
    style?: StyleValue;
  },
): GridBlock {
  return {
    type: 'grid',
    rows: Array.from({ length: rows }).map(() => ({
      cells: [
        {
          value: '',
          style: options?.style,
        },
      ],
    })),
  };
}
