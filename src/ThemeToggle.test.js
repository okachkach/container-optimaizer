import React from 'react';
import { act, cleanup, render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import '@testing-library/jest-dom';
import ThemeToggle from './ThemeToggle';

const STORAGE_KEY = 'container-optimizer-theme';
let mediaQuery;
let listeners;

beforeEach(() => {
  listeners = new Set();
  mediaQuery = {
    matches: false,
    addEventListener: (type, listener) => listeners.add(listener),
    removeEventListener: (type, listener) => listeners.delete(listener),
  };
  Object.defineProperty(window, 'matchMedia', { configurable: true, value: jest.fn(() => mediaQuery) });
});

afterEach(() => {
  cleanup();
  jest.restoreAllMocks();
  window.localStorage.clear();
  document.documentElement.classList.remove('dark');
  document.documentElement.style.colorScheme = '';
  delete window.matchMedia;
});

function changeSystemTheme(dark) {
  act(() => {
    mediaQuery.matches = dark;
    listeners.forEach(listener => listener({ matches: dark }));
  });
}

test.each([false, true])('uses the system preference on first visit (dark: %s)', dark => {
  mediaQuery.matches = dark;
  render(<ThemeToggle />);
  expect(document.documentElement.classList.contains('dark')).toBe(dark);
  expect(document.documentElement.style.colorScheme).toBe(dark ? 'dark' : 'light');
  expect(screen.getByRole('button', { name: `Switch to ${dark ? 'light' : 'dark'} mode` })).toBeEnabled();
  expect(window.localStorage.getItem(STORAGE_KEY)).toBeNull();
});

test('follows system changes until a keyboard selection is made, then persists across remounts', () => {
  const view = render(<ThemeToggle />);
  changeSystemTheme(true);
  expect(document.documentElement).toHaveClass('dark');

  userEvent.tab();
  expect(screen.getByRole('button')).toHaveFocus();
  userEvent.keyboard('{Enter}');
  expect(document.documentElement).not.toHaveClass('dark');
  expect(window.localStorage.getItem(STORAGE_KEY)).toBe('light');

  changeSystemTheme(false);
  changeSystemTheme(true);
  expect(document.documentElement).not.toHaveClass('dark');
  view.unmount();
  expect(listeners.size).toBe(0);
  render(<ThemeToggle />);
  expect(document.documentElement).not.toHaveClass('dark');
});

test.each(['light', 'dark'])('restores an explicit %s preference over the system setting', preference => {
  window.localStorage.setItem(STORAGE_KEY, preference);
  mediaQuery.matches = preference !== 'dark';
  render(<ThemeToggle />);
  expect(document.documentElement.classList.contains('dark')).toBe(preference === 'dark');
  userEvent.click(screen.getByRole('button'));
  const next = preference === 'dark' ? 'light' : 'dark';
  expect(window.localStorage.getItem(STORAGE_KEY)).toBe(next);
  expect(document.documentElement.style.colorScheme).toBe(next);
});

test('uses the system preference for an invalid saved value', () => {
  window.localStorage.setItem(STORAGE_KEY, 'invalid');
  mediaQuery.matches = true;
  render(<ThemeToggle />);
  expect(document.documentElement).toHaveClass('dark');
});

test('still toggles and retains the session choice when storage reads and writes fail', () => {
  jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('Storage blocked'); });
  jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('Storage blocked'); });
  render(<ThemeToggle />);
  userEvent.click(screen.getByRole('button', { name: 'Switch to dark mode' }));
  expect(document.documentElement).toHaveClass('dark');
  changeSystemTheme(true);
  changeSystemTheme(false);
  expect(document.documentElement).toHaveClass('dark');
});

test('syncs theme changes from another tab and resumes system preference when cleared', () => {
  render(<ThemeToggle />);
  act(() => window.dispatchEvent(new StorageEvent('storage', { key: STORAGE_KEY, newValue: 'dark' })));
  expect(document.documentElement).toHaveClass('dark');
  act(() => window.dispatchEvent(new StorageEvent('storage', { key: 'unrelated', newValue: 'light' })));
  expect(document.documentElement).toHaveClass('dark');
  act(() => window.dispatchEvent(new StorageEvent('storage', { key: STORAGE_KEY, newValue: null })));
  expect(document.documentElement).not.toHaveClass('dark');
  changeSystemTheme(true);
  expect(document.documentElement).toHaveClass('dark');
});

test('falls back to light when matchMedia is unavailable', () => {
  delete window.matchMedia;
  render(<ThemeToggle />);
  expect(document.documentElement).not.toHaveClass('dark');
  userEvent.click(screen.getByRole('button'));
  expect(document.documentElement).toHaveClass('dark');
});

test.each([
  ['dark', false, true],
  ['light', true, false],
  [null, true, true],
  [null, false, false],
  ['invalid', true, true],
])('boot script applies %s preference / system dark %s before React mounts', (preference, systemDark, expectedDark) => {
  if (preference) window.localStorage.setItem(STORAGE_KEY, preference);
  mediaQuery.matches = systemDark;
  jest.isolateModules(() => { require('../public/theme'); });
  expect(document.documentElement.classList.contains('dark')).toBe(expectedDark);
  expect(document.documentElement.style.colorScheme).toBe(expectedDark ? 'dark' : 'light');
});

test('boot script tolerates unavailable browser storage', () => {
  jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('Storage blocked'); });
  mediaQuery.matches = true;
  jest.isolateModules(() => { require('../public/theme'); });
  expect(document.documentElement).toHaveClass('dark');
});
