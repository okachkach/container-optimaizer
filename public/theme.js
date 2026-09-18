// Apply the preference before the app paints. Keep the key in sync with ThemeToggle.
(function () {
  var preference = null;
  try {
    preference = localStorage.getItem('container-optimizer-theme');
  } catch (error) {
    // Theme selection still works when browser storage is unavailable.
  }
  var dark = preference === 'dark' ||
    (preference !== 'light' && Boolean(window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches));
  document.documentElement.classList.toggle('dark', dark);
  document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
}());
