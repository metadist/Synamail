import { describe, expect, it } from 'vitest'
import { renderMarkdown } from '@shared/markdown'

describe('renderMarkdown', () => {
  it('renders bold, italic and inline code', () => {
    expect(renderMarkdown('a **b** c')).toContain('<strong>b</strong>')
    expect(renderMarkdown('a *b* c')).toContain('<em>b</em>')
    expect(renderMarkdown('use `npm run dev`')).toContain('<code>npm run dev</code>')
  })

  it('renders the real chat example with bold numbers', () => {
    const html = renderMarkdown('Die Gesamtlänge beträgt **21.196 Kilometer**.')
    expect(html).toContain('<strong>21.196 Kilometer</strong>')
    expect(html.startsWith('<p>')).toBe(true)
  })

  it('renders unordered and ordered lists', () => {
    expect(renderMarkdown('- one\n- two')).toBe('<ul>\n<li>one</li>\n<li>two</li>\n</ul>')
    expect(renderMarkdown('1. one\n2. two')).toBe('<ol>\n<li>one</li>\n<li>two</li>\n</ol>')
  })

  it('strips dangerous HTML instead of showing it as text', () => {
    const html = renderMarkdown('<img src=x onerror=alert(1)> **safe**')
    expect(html).not.toContain('<img')
    expect(html).not.toContain('&lt;img')
    expect(html).not.toContain('onerror')
    expect(html).toContain('<strong>safe</strong>')
  })

  it('removes <script>/<style> blocks entirely', () => {
    const html = renderMarkdown('Hello<script>alert(1)</script> world')
    expect(html).not.toContain('script')
    expect(html).not.toContain('alert')
    expect(html).toContain('Hello')
    expect(html).toContain('world')
  })

  it('converts common formatting HTML into rendered Markdown', () => {
    expect(renderMarkdown('<strong>bold</strong>')).toContain('<strong>bold</strong>')
    expect(renderMarkdown('<em>italic</em>')).toContain('<em>italic</em>')
    expect(renderMarkdown('a<br>b')).toContain('a<br>b')
  })

  it('renders an HTML bullet list as a real list', () => {
    const html = renderMarkdown('<ul><li>one</li><li>two</li></ul>')
    expect(html).toContain('<ul>')
    expect(html).toContain('<li>one</li>')
    expect(html).toContain('<li>two</li>')
    expect(html).not.toContain('&lt;')
  })

  it('decodes basic HTML entities once', () => {
    expect(renderMarkdown('Tom &amp; Jerry')).toContain('Tom &amp; Jerry')
    expect(renderMarkdown('Tom &amp; Jerry')).not.toContain('&amp;amp;')
  })

  it('only allows safe link schemes', () => {
    expect(renderMarkdown('[ok](https://example.com)')).toContain(
      '<a href="https://example.com" target="_blank" rel="noopener noreferrer">ok</a>',
    )
    // javascript: scheme is dropped, label retained, no anchor emitted.
    const bad = renderMarkdown('[x](javascript:alert(1))')
    expect(bad).not.toContain('<a ')
    expect(bad).toContain('x')
  })

  it('renders fenced code blocks without inline parsing inside', () => {
    const html = renderMarkdown('```\n**not bold**\n```')
    expect(html).toContain('<pre class="md-pre"><code>**not bold**</code></pre>')
  })
})
