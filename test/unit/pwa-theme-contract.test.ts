import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const css = readFileSync(resolve(process.cwd(), 'src/styles/base.css'), 'utf8');

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
  });

  it('keeps keyboard focus and disabled states visible', () => {
    expect(css).toContain('outline: 2px solid var(--focus);');
    expect(css).toContain('cursor: not-allowed;');
    expect(css).toContain('opacity: 0.5;');
  });
});
