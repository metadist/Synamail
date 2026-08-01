/**
 * Human-readable Synaplan chat titles for Outlook-sourced actions.
 * Synaplan's History list truncates around 40–60 characters; we cap here
 * so the add-in sends titles that already fit.
 */

const TITLE_MAX = 60

/** Collapse whitespace and truncate with an ellipsis when over the cap. */
export function truncateTitle(text: string, max = TITLE_MAX): string {
  const cleaned = text.replace(/\s+/g, ' ').trim()
  if (cleaned.length <= max) return cleaned
  return `${cleaned.slice(0, max - 1).trimEnd()}…`
}

/**
 * Prefer `Prefix: subject` when a subject exists; otherwise the prefix alone
 * (or a fallback such as a truncated user intent / question).
 */
export function titleWithSubject(prefix: string, subject?: string, fallback?: string): string {
  const sub = subject?.trim()
  if (sub) return truncateTitle(`${prefix}: ${sub}`)
  const alt = fallback?.trim()
  if (alt) return truncateTitle(alt)
  return truncateTitle(prefix)
}
