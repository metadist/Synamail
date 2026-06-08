/**
 * Minimal, dependency-free, XSS-safe Markdown renderer for chat/AI output.
 *
 * Strategy:
 *   1. Assistant replies sometimes arrive as HTML instead of Markdown (or a
 *      mix). `htmlToMarkdownish` converts the handful of common formatting tags
 *      to their Markdown equivalents and STRIPS every other tag, so raw markup
 *      (`<p>`, `<script>`, `<img onerror>`, …) never leaks through as visible
 *      text or executable HTML.
 *   2. Decode the basic HTML entities that conversion may leave behind, so they
 *      don't get double-escaped in the next step.
 *   3. Escape ALL remaining `<`/`>`/`&`, then layer on a curated subset of
 *      Markdown by inserting only known-safe tags. Because the source is fully
 *      escaped before any tag is emitted, model/user content can never inject
 *      markup; the only attribute we emit is a link `href`, whose scheme is
 *      allow-listed (http/https/mailto).
 *
 * Supported: fenced + inline code, bold, italic, links, headings, and
 * unordered/ordered lists, with blank-line paragraphs and single-newline
 * <br>. Not a full CommonMark parser — good enough for assistant replies.
 */

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * Fold a small set of common HTML formatting tags into Markdown tokens and
 * remove everything else. Runs only when the input actually looks like it
 * contains tags. The result is plain text + Markdown — never HTML — which the
 * escaping renderer below then turns into a safe, fixed tag set.
 */
function htmlToMarkdownish(s: string): string {
  if (!/<[a-z!/][^>]*>/i.test(s)) return s
  return (
    s
      // Drop script/style outright, including their contents.
      .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, '')
      // Angle-bracket autolinks → the bare URL, so they don't get stripped to
      // nothing (or left as a dangling `<`) by the catch-all tag removal below.
      .replace(/<((?:https?:\/\/|mailto:)[^>\s]+)>/gi, '$1')
      // Line breaks and block boundaries become newlines.
      .replace(/<\s*br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|h[1-6]|tr|table|blockquote|pre|ul|ol)\s*>/gi, '\n')
      // List items become Markdown bullets.
      .replace(/<li\b[^>]*>/gi, '\n- ')
      // Emphasis maps to Markdown (open and close map to the same token).
      .replace(/<\/?(strong|b)\b[^>]*>/gi, '**')
      .replace(/<\/?(em|i)\b[^>]*>/gi, '*')
      // Strip every remaining tag.
      .replace(/<[^>]*>/g, '')
      // Collapse the runs of blank lines that block-tag conversion can create.
      .replace(/\n{3,}/g, '\n\n')
  )
}

/**
 * Decode the handful of entities that real HTML carries, so they render as the
 * intended character instead of being re-escaped into a literal `&amp;…`.
 * `&amp;` is decoded last so encoded sequences like `&amp;lt;` don't collapse
 * into `<`.
 */
function decodeEntities(s: string): string {
  return s
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#0*39;|&apos;/gi, "'")
    .replace(/&amp;/gi, '&')
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
}

function safeHref(url: string): string | null {
  // `url` may still contain the escaped `&amp;`; decode it for scheme checks
  // and the final href, then attribute-escape.
  const decoded = url.replace(/&amp;/g, '&').trim()
  if (/^(https?:\/\/|mailto:)/i.test(decoded)) return decoded
  return null
}

function renderInline(text: string): string {
  let out = text

  // Inline code first so its contents aren't treated as emphasis.
  out = out.replace(/`([^`]+)`/g, (_m, c: string) => `<code>${c}</code>`)

  // Links: [label](url) — drop the link (keep label) when the scheme is unsafe.
  out = out.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_m, label: string, url: string) => {
    const href = safeHref(url)
    if (!href) return label
    return `<a href="${escapeAttr(href)}" target="_blank" rel="noopener noreferrer">${label}</a>`
  })

  // Bold before italic so `**x**` isn't half-consumed by the italic rule.
  out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
  out = out.replace(/__([^_]+)__/g, '<strong>$1</strong>')
  out = out.replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>')
  out = out.replace(/(^|[^_])_([^_\n]+)_/g, '$1<em>$2</em>')

  return out
}

export function renderMarkdown(src: string): string {
  if (!src) return ''
  const normalized = decodeEntities(htmlToMarkdownish(src)).replace(/\r\n/g, '\n')
  const lines = escapeHtml(normalized).split('\n')
  const html: string[] = []
  let listType: 'ul' | 'ol' | null = null
  let paragraph: string[] = []

  const closeList = (): void => {
    if (listType) {
      html.push(`</${listType}>`)
      listType = null
    }
  }
  const flushPara = (): void => {
    if (paragraph.length) {
      html.push(`<p>${paragraph.map(renderInline).join('<br>')}</p>`)
      paragraph = []
    }
  }

  let i = 0
  while (i < lines.length) {
    const line = lines[i]

    const fence = line.match(/^```/)
    if (fence) {
      flushPara()
      closeList()
      const code: string[] = []
      i++
      while (i < lines.length && !/^```/.test(lines[i])) {
        code.push(lines[i])
        i++
      }
      i++ // skip the closing fence
      html.push(`<pre class="md-pre"><code>${code.join('\n')}</code></pre>`)
      continue
    }

    if (/^\s*$/.test(line)) {
      flushPara()
      closeList()
      i++
      continue
    }

    const heading = line.match(/^(#{1,6})\s+(.*)$/)
    if (heading) {
      flushPara()
      closeList()
      html.push(`<strong class="md-h">${renderInline(heading[2])}</strong>`)
      i++
      continue
    }

    const unordered = line.match(/^\s*[-*+]\s+(.*)$/)
    if (unordered) {
      flushPara()
      if (listType !== 'ul') {
        closeList()
        html.push('<ul>')
        listType = 'ul'
      }
      html.push(`<li>${renderInline(unordered[1])}</li>`)
      i++
      continue
    }

    const ordered = line.match(/^\s*\d+[.)]\s+(.*)$/)
    if (ordered) {
      flushPara()
      if (listType !== 'ol') {
        closeList()
        html.push('<ol>')
        listType = 'ol'
      }
      html.push(`<li>${renderInline(ordered[1])}</li>`)
      i++
      continue
    }

    closeList()
    paragraph.push(line)
    i++
  }

  flushPara()
  closeList()
  return html.join('\n')
}
