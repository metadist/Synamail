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

  it('escapes HTML to prevent injection', () => {
    const html = renderMarkdown('<img src=x onerror=alert(1)> **safe**')
    expect(html).not.toContain('<img')
    expect(html).toContain('&lt;img')
    expect(html).toContain('<strong>safe</strong>')
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
