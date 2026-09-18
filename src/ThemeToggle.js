import { useEffect, useState } from 'react';
import { Moon, Sun } from 'lucide-react';

const STORAGE_KEY = 'container-optimizer-theme';
const SYSTEM_THEME = '(prefers-color-scheme: dark)';
const validPreference = value => value === 'light' || value === 'dark' ? value : null;

function readPreference() {
  try {
    return validPreference(window.localStorage.getItem(STORAGE_KEY));
  } catch {
    return null;
  }
}

export default function ThemeToggle() {
  const [preference, setPreference] = useState(readPreference);
  const [systemDark, setSystemDark] = useState(() => Boolean(window.matchMedia?.(SYSTEM_THEME).matches));
  const dark = preference ? preference === 'dark' : systemDark;

  useEffect(() => {
    const query = window.matchMedia?.(SYSTEM_THEME);
    if (!query) return undefined;
    const update = event => setSystemDark(event.matches);
    setSystemDark(query.matches);
    query.addEventListener('change', update);
    return () => query.removeEventListener('change', update);
  }, []);

  useEffect(() => {
    const update = event => {
      if (event.key === STORAGE_KEY || event.key === null) {
        setPreference(validPreference(event.newValue));
      }
    };
    window.addEventListener('storage', update);
    return () => window.removeEventListener('storage', update);
  }, []);

  useEffect(() => {
    document.documentElement.classList.toggle('dark', dark);
    document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
  }, [dark]);

  const toggle = () => {
    const next = dark ? 'light' : 'dark';
    setPreference(next);
    try {
      window.localStorage.setItem(STORAGE_KEY, next);
    } catch {
      // Keep the user's choice for this session even if it cannot be saved.
    }
  };
  const label = `Switch to ${dark ? 'light' : 'dark'} mode`;
  const Icon = dark ? Sun : Moon;

  return (
    <button
      type="button"
      onClick={toggle}
      aria-label={label}
      title={label}
      className="inline-flex h-10 w-10 items-center justify-center rounded-full border border-slate-200 bg-white text-slate-600 shadow-sm transition-colors hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300 dark:hover:bg-slate-800"
    >
      <Icon className="h-5 w-5" aria-hidden="true" />
    </button>
  );
}
