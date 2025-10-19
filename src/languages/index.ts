import { en } from './en';
import { hu } from './hu';

export const languages = {
  en,
  hu
};

export type Language = keyof typeof languages;
export type LanguageKeys = typeof en;
