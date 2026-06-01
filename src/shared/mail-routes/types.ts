/**
 * Mail Routes — Synamail's per-email AI automations.
 *
 * A Mail Route is "WHEN a mail matches conditions THEN run AI/agent actions on
 * it" (see docs/MAIL_ROUTES.md). v1 models routes as a small set of typed
 * *kinds* rather than the fully generic condition/action engine: each kind maps
 * directly to one of the user-defined automations and carries only the config
 * that automation needs. The discriminated union keeps `vue-tsc --strict`
 * honest and the config UI simple. New kinds extend the union without breaking
 * existing stored routes.
 *
 * Routes are stored per-mailbox in `roamingSettings.routes` — no Synaplan-side
 * schema change. See `useRoamingSettings` for persistence.
 */

export type MailRouteKind = 'meeting' | 'projectIngest' | 'newsletterKb'

export interface MailRouteBase {
  /** Stable client-generated id (used as the Vue key and the edit target). */
  id: string
  /** Human label shown in the config list, e.g. "Oliver — meetings". */
  name: string
  /** Routes are opt-in; a disabled route is never evaluated. */
  enabled: boolean
  kind: MailRouteKind
  /**
   * Sender patterns this route applies to. Each entry is either a full address
   * (`oliver@acme.com`) or a domain (`@acme.com` / `acme.com`). Matching is
   * case-insensitive — see `senderMatches`.
   */
  senders: string[]
}

/**
 * Rule 1 — ICS / date helper. For matching senders, the AI decides whether the
 * mail proposes a meeting time; the engine checks the calendar for a conflict
 * and either re-forwards the request ("you already have a meeting with X at Y")
 * or offers to confirm with a generated `.ics`.
 */
export interface MeetingRoute extends MailRouteBase {
  kind: 'meeting'
  /** Fallback meeting length (minutes) when the mail gives no explicit end. */
  durationMinutes: number
}

/**
 * Shared shape for the two knowledge-base ingestion kinds (rules 2 & 3).
 * Content is translated into `workingLanguage` before it is added so the KB is
 * consistent and searchable in one language; the original is kept too, so
 * search results can reveal the source text in its original language.
 */
export interface IngestRouteBase extends MailRouteBase {
  /** RAG group key the captured content is added to (the "project"). */
  groupId: string
  /** Tag stamped on every ingested artifact (BFILES metadata) for filtering. */
  tag: string
  /** Working-language code (e.g. `en`). Off-language content is translated. */
  workingLanguage: string
  /** Body length (chars) above which the mail text itself is worth ingesting. */
  minBodyChars: number
  /** Also ingest allow-listed document attachments (PDF, DOCX, …). */
  includeAttachments: boolean
}

export interface ProjectIngestRoute extends IngestRouteBase {
  kind: 'projectIngest'
}

export interface NewsletterKbRoute extends IngestRouteBase {
  kind: 'newsletterKb'
}

export type MailRoute = MeetingRoute | ProjectIngestRoute | NewsletterKbRoute

/** Persisted Mail Routes state (per mailbox, in roaming settings). */
export interface MailRoutesState {
  /** Global kill-switch — when true, no route runs. */
  paused: boolean
  routes: MailRoute[]
  /**
   * `internetMessageId` → epoch-ms of last auto-run, so an auto action (e.g.
   * "add to knowledge") runs at most once per message even across reloads.
   */
  seen: Record<string, number>
}

export const EMPTY_ROUTES_STATE: MailRoutesState = {
  paused: false,
  routes: [],
  seen: {},
}

/** Type guard narrowing a route to one of the two ingestion kinds. */
export function isIngestRoute(route: MailRoute): route is ProjectIngestRoute | NewsletterKbRoute {
  return route.kind === 'projectIngest' || route.kind === 'newsletterKb'
}
