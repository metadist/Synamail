/**
 * Synamail plugin — Synaplan web UI.
 *
 * The profiles themselves are created and updated from the Synamail Outlook
 * add-in; this panel is the transparency/privacy surface inside Synaplan:
 * browse every stored contact profile, search them, and delete any of them.
 *
 * The panel is a self-contained vanilla-JS module. It styles itself with the
 * Synaplan design tokens (var(--brand), var(--bg-chip), var(--txt-*), …) that
 * live on :root/.dark, so it matches the app in both light and dark mode
 * without pulling in the Vue/Tailwind build.
 *
 * Mount contract (see synaplan PluginView.vue):
 *   export default { mount(el, { userId, apiBaseUrl, pluginBaseUrl }) }
 */

const FALLBACK_LANG = 'en'
const DOCS_URL = 'https://docs.synaplan.com/synamail'
const REPO_URL = 'https://github.com/metadist/Synamail'
const PAGE_SIZE = 30

const state = {
  el: null,
  ctx: null,
  t: (key) => key,
  profiles: [],
  filter: '',
  page: 1,
  expanded: new Set(),
  listEl: null,
  pagerEl: null,
  countEl: null,
  loadError: null,
}

async function loadI18n(pluginBaseUrl) {
  const lang = (document.documentElement.lang || navigator.language || FALLBACK_LANG)
    .slice(0, 2)
    .toLowerCase()
  const load = async (code) => {
    const res = await fetch(`${pluginBaseUrl}/i18n/${code}.json?v=${Date.now()}`, {
      credentials: 'include',
    })
    if (!res.ok) throw new Error(`i18n ${code} unavailable`)
    return res.json()
  }
  let messages
  try {
    messages = await load(lang)
  } catch {
    messages = await load(FALLBACK_LANG).catch(() => ({}))
  }
  state.t = (key) => messages[key] ?? key
}

function api(path, options = {}) {
  const { userId, apiBaseUrl } = state.ctx
  return fetch(`${apiBaseUrl}/api/v1/user/${userId}/plugins/synamail${path}`, {
    credentials: 'include',
    headers: { Accept: 'application/json', ...(options.headers ?? {}) },
    ...options,
  })
}

function esc(s) {
  const div = document.createElement('div')
  div.textContent = String(s ?? '')
  return div.innerHTML
}

function fmtDate(iso) {
  if (!iso) return '—'
  const d = new Date(iso)
  return Number.isNaN(d.getTime())
    ? '—'
    : d.toLocaleDateString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
      })
}

function initials(profile) {
  const src = (profile.name || profile.email || '?').trim()
  const parts = src.split(/[\s@._-]+/).filter(Boolean)
  const chars = (parts[0]?.[0] ?? '') + (parts[1]?.[0] ?? '')
  return (chars || src[0] || '?').toUpperCase()
}

function matchesFilter(profile, needle) {
  if (!needle) return true
  const haystack = [profile.name, profile.email, profile.org, profile.summary]
    .filter(Boolean)
    .join(' ')
    .toLowerCase()
  return haystack.includes(needle)
}

/** Windowed page list with ellipsis gaps, e.g. [1, '…', 4, 5, 6, '…', 12]. */
function pageWindow(current, total) {
  if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1)
  const set = new Set([1, 2, total - 1, total, current - 1, current, current + 1])
  const pages = [...set].filter((p) => p >= 1 && p <= total).sort((a, b) => a - b)
  const out = []
  let prev = 0
  for (const p of pages) {
    if (prev && p - prev > 1) out.push('…')
    out.push(p)
    prev = p
  }
  return out
}

async function deleteProfile(email) {
  const { t } = state
  if (!window.confirm(`${t('confirmDelete')} (${email})`)) return
  const res = await api(`/profiles/${encodeURIComponent(email)}`, { method: 'DELETE' })
  if (res.ok) {
    state.profiles = state.profiles.filter((p) => p.email !== email)
    state.expanded.delete(email)
    renderList()
  }
}

function renderLoops(loops) {
  const { t } = state
  if (!Array.isArray(loops) || loops.length === 0) return ''
  const items = loops
    .map((loop) => {
      const isMine = /^me\s*:/i.test(loop)
      const isTheirs = /^them\s*:/i.test(loop)
      const text = loop.replace(/^(me|them)\s*:\s*/i, '')
      const badge = isMine
        ? `<span class="sm-owner sm-owner--you">${esc(t('owedByYou'))}</span>`
        : isTheirs
          ? `<span class="sm-owner sm-owner--them">${esc(t('owedByThem'))}</span>`
          : ''
      return `<li>${badge}<span>${esc(text)}</span></li>`
    })
    .join('')
  return `
    <div class="sm-section">
      <div class="sm-section-label">${esc(t('openLoops'))}</div>
      <ul class="sm-loops">${items}</ul>
    </div>`
}

function renderFacts(facts) {
  const { t } = state
  if (!Array.isArray(facts) || facts.length === 0) return ''
  const items = facts.map((f) => `<li>${esc(f)}</li>`).join('')
  return `
    <div class="sm-section">
      <div class="sm-section-label">${esc(t('facts'))}</div>
      <ul class="sm-facts">${items}</ul>
    </div>`
}

function renderPanel(p) {
  const { t } = state
  const tone = p.tone ? `<span class="sm-tone">${esc(p.tone)}</span>` : ''
  const meta = [
    [t('organization'), p.org || '—'],
    [t('firstSeen'), fmtDate(p.firstSeen)],
    [t('lastInbound'), fmtDate(p.lastInbound)],
    [t('lastOutbound'), fmtDate(p.lastOutbound)],
  ]
    .map(
      ([label, value]) =>
        `<div class="sm-meta-cell"><dt>${esc(label)}</dt><dd>${esc(value)}</dd></div>`,
    )
    .join('')

  return `
    <div class="sm-panel">
      ${tone}
      ${p.summary ? `<p class="sm-summary">${esc(p.summary)}</p>` : ''}
      ${renderFacts(p.facts)}
      ${renderLoops(p.openLoops)}
      <dl class="sm-meta-grid">${meta}</dl>
      <div class="sm-panel-actions">
        <button type="button" class="sm-btn sm-btn--danger" data-email="${esc(p.email)}">
          ${esc(t('delete'))}
        </button>
      </div>
    </div>`
}

function renderList() {
  const { t, listEl, pagerEl, countEl } = state

  if (state.loadError) {
    listEl.innerHTML = `<p class="sm-error">${esc(state.loadError)}</p>`
    pagerEl.innerHTML = ''
    countEl.textContent = ''
    return
  }

  const total = state.profiles.length
  if (total === 0) {
    listEl.innerHTML = `<p class="sm-empty">${esc(t('empty'))}</p>`
    pagerEl.innerHTML = ''
    countEl.textContent = ''
    return
  }

  const filtered = state.profiles.filter((p) => matchesFilter(p, state.filter))
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE))
  if (state.page > totalPages) state.page = totalPages
  const start = (state.page - 1) * PAGE_SIZE
  const pageItems = filtered.slice(start, start + PAGE_SIZE)

  countEl.textContent = `${filtered.length} ${t('profiles')}`

  if (filtered.length === 0) {
    listEl.innerHTML = `<p class="sm-empty">${esc(t('noResults'))}</p>`
    pagerEl.innerHTML = ''
    return
  }

  listEl.innerHTML = pageItems
    .map((p) => {
      const open = state.expanded.has(p.email)
      return `
        <div class="sm-item${open ? ' sm-item--open' : ''}">
          <button type="button" class="sm-acc" data-toggle="${esc(p.email)}" aria-expanded="${open}">
            <span class="sm-avatar" aria-hidden="true">${esc(initials(p))}</span>
            <span class="sm-acc-main">
              <span class="sm-acc-name">${esc(p.name || p.email)}</span>
              <span class="sm-acc-sub">${esc(p.email)}${p.org ? ` · ${esc(p.org)}` : ''}</span>
            </span>
            <span class="sm-acc-meta">
              <span>${esc(String(p.emailCount ?? 0))} ${esc(t('emails'))}</span>
              <span class="sm-acc-updated">${esc(t('updated'))} ${esc(fmtDate(p.updatedAt))}</span>
            </span>
            <svg class="sm-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
              <path d="M6 9l6 6 6-6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
          </button>
          ${open ? renderPanel(p) : ''}
        </div>`
    })
    .join('')

  renderPager(totalPages)
  bindListEvents()
}

function renderPager(totalPages) {
  const { t, pagerEl } = state
  if (totalPages <= 1) {
    pagerEl.innerHTML = ''
    return
  }
  const btn = (label, page, opts = {}) => {
    const { disabled = false, active = false, isGap = false } = opts
    if (isGap) return `<span class="sm-page-gap">…</span>`
    const cls = `sm-page-btn${active ? ' sm-page-btn--active' : ''}`
    return `<button type="button" class="${cls}" data-page="${page}" ${disabled ? 'disabled' : ''}>${esc(label)}</button>`
  }

  const numbers = pageWindow(state.page, totalPages)
    .map((p) =>
      p === '…' ? btn('', 0, { isGap: true }) : btn(String(p), p, { active: p === state.page }),
    )
    .join('')

  pagerEl.innerHTML = `
    ${btn(t('prev'), state.page - 1, { disabled: state.page <= 1 })}
    <span class="sm-page-nums">${numbers}</span>
    ${btn(t('next'), state.page + 1, { disabled: state.page >= totalPages })}`

  for (const el of pagerEl.querySelectorAll('button[data-page]')) {
    el.addEventListener('click', () => {
      const target = Number(el.getAttribute('data-page'))
      if (!target || target === state.page) return
      state.page = target
      renderList()
      state.listEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' })
    })
  }
}

function bindListEvents() {
  const { listEl } = state
  for (const el of listEl.querySelectorAll('button[data-toggle]')) {
    el.addEventListener('click', () => {
      const email = el.getAttribute('data-toggle')
      if (state.expanded.has(email)) state.expanded.delete(email)
      else state.expanded.add(email)
      renderList()
    })
  }
  for (const el of listEl.querySelectorAll('button[data-email]')) {
    el.addEventListener('click', (e) => {
      e.stopPropagation()
      deleteProfile(el.getAttribute('data-email'))
    })
  }
}

function shellStyles() {
  return `
  <style>
    .sm-wrap { color: var(--txt-primary); }
    .sm-header { margin: 0 0 20px; }
    .sm-title-row { display: flex; align-items: center; gap: 12px; }
    .sm-title-icon {
      display: inline-flex; align-items: center; justify-content: center;
      width: 40px; height: 40px; border-radius: 10px;
      background: var(--brand-alpha-light); color: var(--brand); flex: none;
    }
    .sm-title { margin: 0; font-size: 1.5rem; font-weight: 600; line-height: 1.2; }
    .sm-tagline { margin: 10px 0 0; color: var(--txt-secondary); max-width: 60ch; }
    .sm-links { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }
    .sm-link-btn {
      display: inline-flex; align-items: center; gap: 7px;
      padding: 8px 14px; border-radius: 9px; font-size: 0.875rem; font-weight: 500;
      text-decoration: none; border: 1px solid var(--border-light);
      transition: background 0.15s ease, color 0.15s ease;
    }
    .sm-link-btn svg { width: 15px; height: 15px; }
    .sm-link-btn--primary { background: var(--brand); color: #fff; border-color: var(--brand); }
    .sm-link-btn--primary:hover { background: var(--brand-hover); }
    .sm-link-btn--ghost { background: var(--brand-alpha-light); color: var(--brand); }
    .sm-link-btn--ghost:hover { background: var(--brand); color: #fff; border-color: var(--brand); }

    .sm-toolbar {
      display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
      margin: 0 0 16px; padding-top: 16px; border-top: 1px solid var(--border-light);
    }
    .sm-search-box { position: relative; flex: 1 1 260px; }
    .sm-search-box svg {
      position: absolute; left: 12px; top: 50%; transform: translateY(-50%);
      width: 16px; height: 16px; color: var(--txt-secondary); pointer-events: none;
    }
    .sm-search {
      width: 100%; padding: 9px 12px 9px 36px; border-radius: 9px;
      border: 1px solid var(--border-light); background: var(--bg-chip);
      color: var(--txt-primary); font-size: 0.9rem; font-family: inherit; box-sizing: border-box;
    }
    .sm-search::placeholder { color: var(--txt-secondary); }
    .sm-search:focus { outline: 2px solid var(--brand); outline-offset: 1px; border-color: transparent; }
    .sm-count { color: var(--txt-secondary); font-size: 0.85rem; white-space: nowrap; }

    .sm-list { display: flex; flex-direction: column; gap: 10px; }
    .sm-item {
      border: 1px solid var(--border-light); border-radius: 12px; overflow: hidden;
      background: var(--bg-chip); transition: box-shadow 0.15s ease;
    }
    .sm-item--open { box-shadow: inset 0 0 0 1px var(--brand-alpha-light); }
    .sm-acc {
      width: 100%; display: flex; align-items: center; gap: 12px;
      padding: 12px 14px; background: transparent; border: 0; cursor: pointer;
      text-align: left; color: var(--txt-primary); font-family: inherit;
    }
    .sm-acc:hover { background: var(--brand-alpha-light); }
    .sm-avatar {
      display: inline-flex; align-items: center; justify-content: center;
      width: 34px; height: 34px; border-radius: 50%; flex: none;
      background: var(--brand); color: #fff; font-size: 0.8rem; font-weight: 600;
    }
    .sm-acc-main { display: flex; flex-direction: column; gap: 2px; min-width: 0; flex: 1; }
    .sm-acc-name { font-weight: 600; font-size: 0.95rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .sm-acc-sub { color: var(--txt-secondary); font-size: 0.8rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .sm-acc-meta { display: flex; flex-direction: column; align-items: flex-end; gap: 2px; font-size: 0.75rem; color: var(--txt-secondary); flex: none; }
    .sm-chevron { color: var(--txt-secondary); flex: none; transition: transform 0.2s ease; }
    .sm-item--open .sm-chevron { transform: rotate(180deg); }

    .sm-panel { padding: 4px 16px 16px; border-top: 1px solid var(--border-light); }
    .sm-tone {
      display: inline-block; margin: 12px 0 0; padding: 3px 12px; border-radius: 999px;
      background: var(--brand-alpha-light); color: var(--brand); font-size: 0.78rem; font-weight: 500;
    }
    .sm-summary { margin: 12px 0 0; line-height: 1.55; white-space: pre-wrap; }
    .sm-section { margin-top: 14px; }
    .sm-section-label {
      font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em;
      color: var(--txt-secondary); margin-bottom: 6px;
    }
    .sm-facts, .sm-loops { margin: 0; padding-left: 18px; display: flex; flex-direction: column; gap: 4px; }
    .sm-loops { list-style: none; padding-left: 0; }
    .sm-loops li { display: flex; align-items: baseline; gap: 8px; }
    .sm-owner {
      flex: none; font-size: 0.68rem; font-weight: 600; padding: 1px 7px; border-radius: 999px;
      text-transform: uppercase; letter-spacing: 0.03em;
    }
    .sm-owner--you { background: var(--status-warning-muted); color: var(--status-warning-text); }
    .sm-owner--them { background: var(--status-success-muted); color: var(--status-success-text); }

    .sm-meta-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
      gap: 10px 16px; margin: 16px 0 0;
    }
    .sm-meta-cell dt { font-size: 0.72rem; color: var(--txt-secondary); text-transform: uppercase; letter-spacing: 0.03em; }
    .sm-meta-cell dd { margin: 2px 0 0; font-size: 0.85rem; }
    .sm-panel-actions { margin-top: 16px; display: flex; justify-content: flex-end; }

    .sm-btn {
      padding: 7px 14px; border-radius: 8px; font-size: 0.85rem; font-weight: 500;
      border: 1px solid var(--border-light); background: transparent; color: var(--txt-primary);
      cursor: pointer; font-family: inherit;
    }
    .sm-btn--danger { color: var(--status-error); border-color: var(--status-error-muted); }
    .sm-btn--danger:hover { background: var(--status-error-muted); }

    .sm-empty, .sm-error { color: var(--txt-secondary); text-align: center; padding: 32px 12px; }
    .sm-error { color: var(--status-error); }

    .sm-pager { display: flex; align-items: center; justify-content: center; gap: 8px; margin-top: 20px; flex-wrap: wrap; }
    .sm-page-nums { display: flex; align-items: center; gap: 4px; }
    .sm-page-btn {
      min-width: 34px; height: 34px; padding: 0 10px; border-radius: 8px;
      border: 1px solid var(--border-light); background: transparent; color: var(--txt-primary);
      cursor: pointer; font-size: 0.85rem; font-family: inherit;
    }
    .sm-page-btn:hover:not(:disabled):not(.sm-page-btn--active) { background: var(--brand-alpha-light); }
    .sm-page-btn--active { background: var(--brand); color: #fff; border-color: var(--brand); cursor: default; }
    .sm-page-btn:disabled { opacity: 0.45; cursor: default; }
    .sm-page-gap { min-width: 20px; text-align: center; color: var(--txt-secondary); }

    @media (max-width: 560px) {
      .sm-acc-meta { display: none; }
    }
  </style>`
}

function buildShell() {
  const { t } = state
  const mailIcon = `<svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 6h16a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1Z" stroke="currentColor" stroke-width="1.6"/><path d="m4 7 8 6 8-6" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  const docsIcon = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M14 3H7a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V8l-4-5Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M14 3v5h4M9 13h6M9 16h6" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  const repoIcon = `<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 2A10 10 0 0 0 8.8 21.5c.5.1.7-.2.7-.5v-1.7c-2.8.6-3.4-1.4-3.4-1.4-.5-1.2-1.1-1.5-1.1-1.5-.9-.6.1-.6.1-.6 1 .1 1.5 1 1.5 1 .9 1.5 2.3 1.1 2.9.8.1-.6.3-1.1.6-1.3-2.2-.3-4.6-1.1-4.6-4.9 0-1.1.4-2 1-2.7-.1-.3-.4-1.3.1-2.7 0 0 .8-.3 2.7 1a9.4 9.4 0 0 1 5 0c1.9-1.3 2.7-1 2.7-1 .5 1.4.2 2.4.1 2.7.6.7 1 1.6 1 2.7 0 3.8-2.4 4.6-4.6 4.9.3.3.6.9.6 1.9v2.8c0 .3.2.6.7.5A10 10 0 0 0 12 2Z"/></svg>`

  state.el.innerHTML = `
    ${shellStyles()}
    <div class="sm-wrap">
      <header class="sm-header">
        <div class="sm-title-row">
          <span class="sm-title-icon">${mailIcon}</span>
          <h2 class="sm-title">${esc(t('title'))}</h2>
        </div>
        <p class="sm-tagline">${esc(t('tagline'))}</p>
        <div class="sm-links">
          <a class="sm-link-btn sm-link-btn--primary" href="${DOCS_URL}" target="_blank" rel="noopener noreferrer">
            ${docsIcon}<span>${esc(t('docsLink'))}</span>
          </a>
          <a class="sm-link-btn sm-link-btn--ghost" href="${REPO_URL}" target="_blank" rel="noopener noreferrer">
            ${repoIcon}<span>${esc(t('repoLink'))}</span>
          </a>
        </div>
      </header>

      <div class="sm-toolbar">
        <div class="sm-search-box">
          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><circle cx="11" cy="11" r="7" stroke="currentColor" stroke-width="2"/><path d="m20 20-3-3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          <input type="search" class="sm-search" placeholder="${esc(t('searchPlaceholder'))}" aria-label="${esc(t('searchPlaceholder'))}" />
        </div>
        <span class="sm-count"></span>
      </div>

      <div class="sm-list"></div>
      <nav class="sm-pager" aria-label="Pagination"></nav>
    </div>`

  state.listEl = state.el.querySelector('.sm-list')
  state.pagerEl = state.el.querySelector('.sm-pager')
  state.countEl = state.el.querySelector('.sm-count')

  const search = state.el.querySelector('.sm-search')
  search.addEventListener('input', () => {
    state.filter = search.value.trim().toLowerCase()
    state.page = 1
    renderList()
  })
}

async function loadProfiles() {
  state.loadError = null
  state.listEl.innerHTML = `<p class="sm-empty">${esc(state.t('loading'))}</p>`
  try {
    const res = await api('/profiles')
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    const data = await res.json()
    state.profiles = data.profiles ?? []
  } catch (err) {
    state.loadError = err instanceof Error ? err.message : String(err)
  }
  renderList()
}

export default {
  async mount(el, context) {
    state.el = el
    state.ctx = context
    state.profiles = []
    state.filter = ''
    state.page = 1
    state.expanded = new Set()
    await loadI18n(context.pluginBaseUrl)
    buildShell()
    await loadProfiles()
  },
}
