export type DayOfWeekKeys = 'MON' | 'TUE' | 'WED' | 'THU' | 'FRI' | 'SAT' | 'SUN';

export type CategoryKeys = 'fried-chicken' | 'variety' | 'curry' | 'udon-soba' | 'ramen' | 'pasta' | 'daily-dinner';

export type CategoryConst = {
  name: string;
  defaultPrice?: string;
}

export type Menu = {
  name: string;
  price: {
    normal: number;
    L?: number;
    M?: number;
    big?: number;
  }
  isSpecialMenu: boolean;
}

export type DailyMenu = Partial<Record<CategoryKeys, Menu>>;