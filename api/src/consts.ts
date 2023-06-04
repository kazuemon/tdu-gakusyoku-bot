import { CategoryConst, CategoryKeys, DayOfWeekKeys } from "./types";

export const dayOfWeekKeys = [
  'MON',
  'TUE',
  'WED',
  'THU',
  'FRI',
  'SAT',
  'SUN',
] satisfies DayOfWeekKeys[];

export const categoryConsts: Record<CategoryKeys, CategoryConst> = {
  'fried-chicken': {
    name: '電大唐揚げ専門店',
  },
  variety: {
    name: 'バラエティー',
  },
  curry: {
    name: 'カレー',
  },
  'udon-soba': {
    name: 'うどん/そば',
    defaultPrice: '￥350 大盛¥400',
  },
  ramen: {
    name: 'ラーメン',
    defaultPrice: '￥350 大盛¥400',
  },
  pasta: {
    name: 'パスタ',
    defaultPrice: '￥400',
  },
  'daily-dinner': {
    name: '日替わり定食',
    defaultPrice: '¥500',
  },
};

export const categories = [
  categoryConsts["fried-chicken"],
  categoryConsts.variety,
  categoryConsts.curry,
  categoryConsts["udon-soba"],
  categoryConsts.ramen,
  categoryConsts.pasta,
  categoryConsts["daily-dinner"],
];