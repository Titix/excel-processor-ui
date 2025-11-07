/**
 * Test file for language files
 */

import { languages } from '@/languages/index';
import { en } from '@/languages/en';
import { hu } from '@/languages/hu';

describe('Languages', () => {
  test('languages object contains en and hu', () => {
    expect(languages).toHaveProperty('en');
    expect(languages).toHaveProperty('hu');
    expect(languages.en).toBe(en);
    expect(languages.hu).toBe(hu);
  });

  test('en language has all required keys', () => {
    expect(en).toHaveProperty('appTitle');
    expect(en).toHaveProperty('appSubtitle');
    expect(en).toHaveProperty('messages');
    expect(en).toHaveProperty('footer');
  });

  test('hu language has all required keys', () => {
    expect(hu).toHaveProperty('appTitle');
    expect(hu).toHaveProperty('appSubtitle');
    expect(hu).toHaveProperty('messages');
    expect(hu).toHaveProperty('footer');
  });

  test('both languages have same structure', () => {
    const enKeys = Object.keys(en);
    const huKeys = Object.keys(hu);
    expect(enKeys.length).toBe(huKeys.length);
    expect(enKeys.sort()).toEqual(huKeys.sort());
  });
});


