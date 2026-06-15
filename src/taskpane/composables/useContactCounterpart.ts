/**
 * Resolve the "contact" a profile should be about from the open email.
 *
 * Profiling is always about the OTHER party, never the mailbox owner. Keying
 * off `from` alone breaks the moment you read your own outbound mail — e.g.
 * after filtering/sorting the list by sender, or browsing Sent Items — because
 * there `from` is you, so the panel would try to profile your own address and
 * "nothing happens". We pick the counterpart instead:
 *
 *   - inbound  (from someone else)  → contact = from
 *   - outbound (from you)           → contact = first recipient that isn't you
 *
 * `direction` is derived from the same decision so the rolled-in email is
 * attributed correctly.
 */

import { computed, type Ref } from 'vue'
import type { OutlookItemSnapshot } from '@/taskpane/composables/useOutlookItem'

function ownerEmail(): string {
  try {
    const addr =
      (typeof Office !== 'undefined' ? Office.context?.mailbox?.userProfile?.emailAddress : '') ??
      ''
    return addr.trim().toLowerCase()
  } catch {
    return ''
  }
}

export function useContactCounterpart(item: Ref<OutlookItemSnapshot>) {
  const owner = computed(ownerEmail)

  const contactEmail = computed<string>(() => {
    if (item.value.mode !== 'read') return ''
    const o = owner.value
    const from = (item.value.from ?? '').trim().toLowerCase()
    if (from && from !== o) return from
    // Outbound (or owner-as-sender): the contact is the first recipient that
    // isn't the mailbox owner; fall back to the first recipient otherwise.
    const recipients = item.value.to.map((a) => a.trim().toLowerCase()).filter(Boolean)
    return recipients.find((a) => a !== o) ?? recipients[0] ?? ''
  })

  const direction = computed<'inbound' | 'outbound'>(() => {
    const from = (item.value.from ?? '').trim().toLowerCase()
    return from && from === contactEmail.value ? 'inbound' : 'outbound'
  })

  return { contactEmail, direction }
}
