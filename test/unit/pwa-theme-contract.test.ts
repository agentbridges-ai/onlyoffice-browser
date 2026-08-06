import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const css = readFileSync(resolve(process.cwd(), 'src/styles/base.css'), 'utf8');
const page = readFileSync(resolve(process.cwd(), 'pages/index.html'), 'utf8');
const manifest = JSON.parse(readFileSync(resolve(process.cwd(), 'public/manifest.webmanifest'), 'utf8')) as {
  background_color: string;
  theme_color: string;
};

describe('standalone PWA theme contract', () => {
  it('shares Piwork light and dark semantic tokens', () => {
    expect(css).toContain('--background: oklch(0.9789 0.0013 106.42);');
    expect(css).toContain('--foreground: oklch(0.3174 0.0091 88.75);');
    expect(css).toContain('--accent: oklch(0.2103 0.0013 106.42);');
    expect(css).toContain('--background: oklch(0.2178 0 0);');
    expect(css).toContain('--foreground: oklch(0.8884 0 0);');
    expect(css).toContain('--accent: oklch(0.8884 0 0);');
  });

  it('uses the Piwork control geometry and semantic surfaces', () => {
    expect(css).toContain('--control-radius: 6px;');
    expect(css).toContain('--panel-radius: 10px;');
    expect(css).toContain('--control-height: 32px;');
    expect(css).toContain('background: var(--surface);');
    expect(css).toContain('background: var(--default);');
    expect(css).toContain('background: var(--accent);');
    expect(css).toContain('height: var(--control-height);');
    expect(css).not.toContain('min-height: 28px;');
    expect(css).not.toContain('min-height: 30px;');
  });

  it('ports the Piwork preview-pane geometry without a second visual token set', () => {
    expect(css).toContain('--piwork-control-radius: var(--control-radius);');
    expect(css).toContain('--piwork-titlebar-height: 40px;');
    expect(css).toContain('--piwork-titlebar-control-size: 31px;');
    expect(css).toContain('.preview-tabbar');
    expect(css).toContain('.document-tab-close');
    expect(css).toContain('.preview-toolbar-button');
    expect(css).toContain('background: var(--accent-soft);');
  });

  it('keeps keyboard focus and disabled states visible', () => {
    expect(css).toContain('outline: 2px solid var(--focus);');
    expect(css).toContain('cursor: not-allowed;');
    expect(css).toContain('opacity: 0.5;');
  });

  it('uses the neutral theme in browser and install surfaces', () => {
    expect(page).toContain('<meta name="theme-color" content="#faf9f7" media="(prefers-color-scheme: light)" />');
    expect(page).toContain('<meta name="theme-color" content="#1c1c1c" media="(prefers-color-scheme: dark)" />');
    expect(manifest).toMatchObject({
      background_color: '#faf9f7',
      theme_color: '#30302e',
    });
  });
});
