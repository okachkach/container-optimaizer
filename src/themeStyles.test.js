import fs from 'fs';
import path from 'path';
import postcss from 'postcss';
import tailwindcss from 'tailwindcss';

let styles;

beforeAll(async () => {
  const css = fs.readFileSync(path.join(__dirname, 'index.css'), 'utf8');
  const result = await postcss([tailwindcss({
    ...require('../tailwind.config'),
    content: [{ raw: fs.readFileSync(path.join(__dirname, 'ThemeToggle.js'), 'utf8'), extension: 'js' }],
  })]).process(css, { from: undefined });
  styles = result.root;
});

afterEach(() => document.documentElement.classList.remove('dark'));

function rootColor(property) {
  let color;
  styles.walkRules(rule => {
    // Only inspect rules that set this color; this excludes placeholder and focus selectors.
    const declarations = rule.nodes.filter(node => node.type === 'decl' && node.prop === property);
    if (declarations.length && document.documentElement.matches(rule.selector)) {
      color = declarations[declarations.length - 1].value;
    }
  });
  return color.match(/^rgb\((\d+) (\d+) (\d+)/).slice(1).map(Number);
}

function luminance(rgb) {
  return rgb.map(channel => {
    const value = channel / 255;
    return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  }).reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index], 0);
}

test.each(['light', 'dark'])('%s mode provides readable inherited text for product details and toolbar buttons', theme => {
  document.documentElement.classList.toggle('dark', theme === 'dark');
  const text = luminance(rootColor('color'));
  const background = luminance(rootColor('background-color'));
  // Also check the panels where the product tables and toolbar inherit this color.
  const panel = luminance(theme === 'dark' ? [15, 23, 42] : [255, 255, 255]);
  for (const surface of [background, panel]) {
    expect((Math.max(text, surface) + 0.05) / (Math.min(text, surface) + 0.05)).toBeGreaterThanOrEqual(4.5);
  }
  expect(theme === 'dark' ? text > background : text < background).toBe(true);
});
