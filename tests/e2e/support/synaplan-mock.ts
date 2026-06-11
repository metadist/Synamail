/**
 * Deterministic Synaplan API mock for the L2 E2E harness.
 *
 * Intercepts every `/api/v1/**` request the real `RealSynaplanClient` makes
 * and answers with the exact wire shapes verified against a live local
 * Synaplan on 2026-06-04 (see docs/PROJECT_PLAN.md). The app's real client code
 * path runs unchanged — real `fetch`, real header/body building, real response
 * parsing — only the network boundary is stubbed, so these specs catch
 * client/UI regressions without needing a backend (and run in CI).
 *
 * To run the SAME specs against a real Synaplan instead, set `SYNAPLAN_E2E_LIVE=1`
 * (plus SYNAPLAN_BASE_URL / SYNAPLAN_API_KEY) and skip calling `mockSynaplan`.
 */

import type { Page, Route } from '@playwright/test'

/** The AI reply text for a `/messages/send`, chosen by the system prompt. */
function aiReplyFor(message: string): string {
  if (message.includes('email triage assistant')) {
    return '{"category":"support","confidence":0.91,"reasoning":"Reports a problem and asks for confirmation."}'
  }
  if (message.includes('extract proposed meeting')) {
    return JSON.stringify([
      {
        title: 'Invoice review',
        start: '2026-06-09T15:00:00',
        end: '2026-06-09T15:30:00',
        location: 'Teams',
      },
    ])
  }
  if (message.includes('professional translator')) {
    return 'Hallo, könnten Sie bitte die Mai-Rechnung bestätigen?'
  }
  if (message.includes('Write a reply in')) {
    return '<p>Thanks Alice — the May invoice #4821 is approved. Friday works.</p>'
  }
  if (message.includes('Summarise the email')) {
    return '- Alice asks to confirm May invoice #4821\n- Payment is due Friday\n- Proposes a review meeting next Tuesday 3pm'
  }
  if (message.includes('answering follow-up questions')) {
    return 'The email asks you to approve invoice #4821 for payment by Friday.'
  }
  if (message.includes('helpful assistant living inside')) {
    return 'In German, "efficiency" is "Effizienz".'
  }
  return 'OK.'
}

function json(route: Route, body: unknown, status = 200): Promise<void> {
  return route.fulfill({
    status,
    contentType: 'application/json',
    body: JSON.stringify(body),
  })
}

/**
 * Register the Synaplan API mock on a page. Call BEFORE `page.goto`.
 * Routes are matched by pathname + method; anything unrecognised gets a
 * generic `{success:true}` so a stray call never hard-fails a flow.
 */
export async function mockSynaplan(page: Page): Promise<void> {
  // Contact AI Profiling (synamail plugin) — stateful per page so the
  // empty-state → update → delete flow behaves like the real plugin.
  const profiles = new Map<string, Record<string, unknown>>()

  await page.route('**/api/v1/**', async (route) => {
    const req = route.request()
    const url = new URL(req.url())
    const path = url.pathname.replace(/^.*\/api\/v1/, '/api/v1')
    const method = req.method()

    // --- identity ---------------------------------------------------------
    if (path === '/api/v1/auth/me' && method === 'GET') {
      return json(route, {
        success: true,
        user: { id: 1, email: 'admin@synaplan.com', level: 'pro', isAdmin: true },
      })
    }

    // --- chat lifecycle ---------------------------------------------------
    if (path === '/api/v1/chats' && method === 'POST') {
      return json(route, {
        success: true,
        chat: {
          id: 101,
          title: 'E2E chat',
          createdAt: '2026-06-04T10:00:00Z',
          updatedAt: '2026-06-04T10:00:00Z',
        },
      })
    }

    // Streaming variant used whenever the UI passes an onChunk handler. The
    // catch-all would answer 200 JSON, which the SSE reader parses as "no
    // frames" → empty result — so a real event-stream reply is required here.
    if (path === '/api/v1/messages/stream' && method === 'GET') {
      const text = aiReplyFor(url.searchParams.get('message') ?? '')
      return route.fulfill({
        status: 200,
        contentType: 'text/event-stream',
        body:
          `data: ${JSON.stringify({ status: 'data', chunk: text })}\n\n` +
          `data: ${JSON.stringify({ status: 'complete' })}\n\n`,
      })
    }

    if (path === '/api/v1/messages/send' && method === 'POST') {
      let message = ''
      try {
        message = (JSON.parse(req.postData() ?? '{}') as { message?: string }).message ?? ''
      } catch {
        /* leave message empty → generic reply */
      }
      return json(route, {
        success: true,
        message: 'sent',
        incomingMessage: { id: 9001 },
        outgoingMessage: { id: 9002, text: aiReplyFor(message), provider: 'mock' },
      })
    }

    // --- RAG / files ------------------------------------------------------
    if (path === '/api/v1/files/groups' && method === 'GET') {
      return json(route, {
        success: true,
        groups: [
          { name: 'general', count: 3 },
          { name: 'project:helios', count: 5 },
        ],
      })
    }

    if (path === '/api/v1/files/upload' && method === 'POST') {
      return json(route, {
        success: true,
        files: [
          {
            success: true,
            id: 555,
            filename: 'message.eml',
            group_key: 'contact:alice@contoso.com',
            vectorized: true,
            chunks_created: 3,
          },
        ],
      })
    }

    if (path === '/api/v1/rag/search' && method === 'POST') {
      return json(route, {
        success: true,
        query: 'invoice',
        total_results: 2,
        results: [
          {
            chunk_id: 'c-1',
            message_id: 42,
            text: 'May invoice #4821 — please confirm approval by Friday.',
            score: 0.86,
            start_line: 1,
            end_line: 3,
          },
          {
            chunk_id: 'c-2',
            message_id: 43,
            text: 'Follow-up: payment scheduled, thanks for confirming.',
            score: 0.74,
            start_line: 1,
            end_line: 2,
          },
        ],
      })
    }

    // --- model config (Settings view) ------------------------------------
    if (path === '/api/v1/config/models/defaults' && method === 'GET') {
      return json(route, { success: true, defaults: { CHAT: 1, TEXT2PIC: 2, VECTORIZE: 3 } })
    }
    if (path === '/api/v1/config/models' && method === 'GET') {
      return json(route, {
        success: true,
        models: {
          CHAT: [{ id: 1, name: 'mock-chat', service: 'mock' }],
          TEXT2PIC: [{ id: 2, name: 'mock-image', service: 'mock' }],
          VECTORIZE: [{ id: 3, name: 'mock-vectorize', service: 'mock' }],
        },
      })
    }

    if (path.startsWith('/api/v1/apikeys/') && method === 'DELETE') {
      return route.fulfill({ status: 204, body: '' })
    }

    // --- Contact AI Profiling (synamail plugin) ---------------------------
    const profileMatch = path.match(
      /^\/api\/v1\/user\/\d+\/plugins\/synamail\/profiles\/([^/]+)(\/update)?$/,
    )
    if (profileMatch) {
      const email = decodeURIComponent(profileMatch[1]).toLowerCase()
      if (method === 'POST' && profileMatch[2] === '/update') {
        const existing = profiles.get(email)
        const count = ((existing?.emailCount as number | undefined) ?? 0) + 1
        const profile = {
          email,
          name: 'Alice Contoso',
          org: 'contoso.com',
          summary: `Alice handles invoicing at Contoso. ${count} email(s) profiled so far; the May invoice #4821 is awaiting confirmation.`,
          tone: 'friendly but business-like',
          facts: ['Works in accounting at Contoso', 'Prefers confirmations by Friday'],
          openLoops: ['me: confirm the May invoice #4821'],
          emailCount: count,
          firstSeen: '2026-06-01T08:00:00+00:00',
          lastInbound: '2026-06-04T09:00:00+00:00',
          updatedAt: '2026-06-04T10:00:00+00:00',
        }
        profiles.set(email, profile)
        return json(route, { success: true, profile })
      }
      if (method === 'DELETE') {
        const deleted = profiles.delete(email)
        return json(route, { success: true, deleted })
      }
      if (method === 'GET') {
        return json(route, { success: true, profile: profiles.get(email) ?? null })
      }
    }

    // --- catch-all --------------------------------------------------------
    return json(route, { success: true })
  })
}
