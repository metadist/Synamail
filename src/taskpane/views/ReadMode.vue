<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import BlockSenderDialog from '@/taskpane/components/BlockSenderDialog.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import SaveToRagDialog from '@/taskpane/components/SaveToRagDialog.vue'
import SenderHistoryList from '@/taskpane/components/SenderHistoryList.vue'
import TonePicker from '@/taskpane/components/TonePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import {
  getChatIdForConversation,
  getLastRagGroupId,
  setChatIdForConversation,
  setLastRagGroupId,
} from '@/taskpane/composables/useRoamingSettings'
import {
  countImageAttachments,
  getImageAttachments,
  getReadItemAsFile,
  useOutlookItem,
} from '@/taskpane/composables/useOutlookItem'
import { useOutlookMailbox } from '@/taskpane/composables/useOutlookMailbox'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { go, openContactKb } from '@/taskpane/router'
import type { MeetingProposal, SenderHistoryResult } from '@shared/types'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()
const mailbox = useOutlookMailbox()

type ActionKey =
  | 'summarise'
  | 'translate'
  | 'reply'
  | 'classify'
  | 'save'
  | 'ask'
  | 'senderHistory'
  | 'senderSummary'
  | 'block'
  | 'meeting'
  | null
const active = ref<ActionKey>(null)
const meetingProposals = ref<MeetingProposal[] | null>(null)
const result = ref<string>('')
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const targetLang = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('en')
const tone = ref<'formal' | 'concise' | 'friendly'>('concise')
const question = ref('')
const askHistory = ref<{ q: string; a: string }[]>([])
const showSaveDialog = ref(false)
const showBlockDialog = ref(false)
const senderHistory = ref<SenderHistoryResult | null>(null)

const senderEmail = computed(() => item.value.from ?? '')
const conversationKey = computed(
  () => item.value.conversationId ?? `synamail:${item.value.subject}`,
)
const imageCount = computed(() => countImageAttachments(item.value))
// Cache the uploaded image file ids so follow-up questions in the same email
// re-reference them (cheap) instead of re-uploading + re-running Vision AI.
const imageFileIds = ref<number[]>([])

async function run<T>(key: NonNullable<ActionKey>, fn: () => Promise<T | null>): Promise<T | null> {
  active.value = key
  error.value = null
  try {
    return await fn()
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
    return null
  } finally {
    active.value = null
  }
}

async function summarise(): Promise<void> {
  const r = await run('summarise', () =>
    call((c) =>
      c.summarise({
        subject: item.value.subject,
        body: item.value.bodyText,
        from: item.value.from,
        to: item.value.to,
      }),
    ),
  )
  if (r) result.value = r.bullets.map((b) => `• ${b}`).join('\n')
}

async function translate(): Promise<void> {
  const r = await run('translate', () =>
    call((c) =>
      c.translate({
        text: item.value.bodyText,
        targetLanguage: targetLang.value === 'auto' ? 'en' : targetLang.value,
      }),
    ),
  )
  if (r) result.value = r.translation
}

async function draftReply(): Promise<void> {
  const r = await run('reply', () =>
    call((c) =>
      c.draftReply({
        subject: item.value.subject,
        body: item.value.bodyText,
        tone: tone.value,
        language: targetLang.value === 'auto' ? 'en' : targetLang.value,
      }),
    ),
  )
  if (r) {
    try {
      Office.context.mailbox.item?.displayReplyForm({ htmlBody: r.htmlBody })
    } catch (err) {
      error.value = err instanceof Error ? err.message : String(err)
    }
  }
}

async function classify(): Promise<void> {
  const r = await run('classify', () =>
    call((c) =>
      c.classify({ subject: item.value.subject, body: item.value.bodyText, from: item.value.from }),
    ),
  )
  if (r) result.value = `${r.category} (${Math.round(r.confidence * 100)}%) — ${r.reasoning}`
}

function openSaveDialog(): void {
  error.value = null
  status.value = null
  showSaveDialog.value = true
}

async function handleSaveConfirm(payload: {
  groupId: string
  processLevel: 'store' | 'extract' | 'vectorize' | 'full'
}): Promise<void> {
  showSaveDialog.value = false
  await run('save', async () => {
    const file = await getReadItemAsFile(item.value)
    const r = await call((c) =>
      c.fileUpload({
        filename: file.filename,
        contentBase64: file.contentBase64,
        mimeType: file.mimeType,
        groupId: payload.groupId,
        metadata: {
          from: item.value.from ?? '',
          subject: item.value.subject,
          to: item.value.to.join(', '),
        },
        processLevel: payload.processLevel,
      }),
    )
    if (r) {
      await setLastRagGroupId(payload.groupId)
      status.value = t('read.saveDialog.savedTo', {
        group: payload.groupId,
        fileId: r.fileId,
      })
    }
    return r
  })
}

async function loadSenderHistory(): Promise<void> {
  const email = senderEmail.value
  if (!email) return
  status.value = null
  senderHistory.value = null
  result.value = ''
  const r = await run('senderHistory', () => mailbox.senderHistory(email, 12))
  if (r) senderHistory.value = r
}

async function summariseSenderHistory(): Promise<void> {
  if (!senderHistory.value) return
  // Build a single body block from the history so the existing summarise
  // prompt produces a multi-message recap. This keeps the AI plumbing
  // unchanged — the new feature is purely an aggregation in the UI layer.
  const block = senderHistory.value.items
    .map((m, i) => `# Message ${i + 1} — ${m.date}\n## ${m.subject}\n${m.snippet}`)
    .join('\n\n')
  const r = await run('senderSummary', () =>
    call((c) =>
      c.summarise({
        subject: t('read.senderHistory.summarySubject', {
          email: senderHistory.value!.email,
          n: senderHistory.value!.total,
        }),
        body: block,
        from: senderHistory.value!.email,
      }),
    ),
  )
  if (r) result.value = r.bullets.map((b) => `• ${b}`).join('\n')
}

function openBlockDialog(): void {
  error.value = null
  status.value = null
  showBlockDialog.value = true
}

async function handleBlockConfirm(payload: { alsoCleanExisting: boolean }): Promise<void> {
  showBlockDialog.value = false
  const email = senderEmail.value
  if (!email) return
  const r = await run('block', () => mailbox.blockSender(email, payload.alsoCleanExisting))
  if (r) {
    status.value = r.serverSide
      ? t('read.blockDialog.successServer', { email, moved: r.movedCount })
      : t('read.blockDialog.successMock', { email })
  }
}

/**
 * Upload the email's images to Synaplan (once) so the AI can "see" them.
 * Synaplan's Vision AI extracts text/description on upload; we then attach the
 * resulting file ids to the chat turn. Cached so follow-ups don't re-upload.
 */
async function ensureImageFileIds(): Promise<number[]> {
  if (imageFileIds.value.length > 0 || imageCount.value === 0) return imageFileIds.value
  const images = await getImageAttachments(item.value)
  const ids: number[] = []
  for (const img of images) {
    const r = await call((c) =>
      c.fileUpload({
        filename: img.filename,
        contentBase64: img.contentBase64,
        mimeType: img.mimeType,
        processLevel: 'extract',
        metadata: { source: 'outlook-email-image', subject: item.value.subject },
      }),
    )
    if (r?.fileId) ids.push(r.fileId)
  }
  imageFileIds.value = ids
  return ids
}

async function ask(): Promise<void> {
  const q = question.value.trim()
  if (!q) return
  const conversationId = conversationKey.value
  const chatId = getChatIdForConversation(conversationId)
  const r = await run('ask', async () => {
    const fileIds = await ensureImageFileIds()
    return call((c) =>
      c.ask({
        conversationId,
        question: q,
        emailContext: item.value.bodyText,
        chatId,
        fileIds: fileIds.length ? fileIds : undefined,
      }),
    )
  })
  if (r) {
    if (!chatId && r.chatId) {
      try {
        await setChatIdForConversation(conversationId, r.chatId)
      } catch {
        // Roaming write can fail in tests / offline; the in-memory chatId
        // still works for the rest of this session.
      }
    }
    askHistory.value.push({ q, a: r.answer })
    question.value = ''
  }
}

function userTimezone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC'
  } catch {
    return 'UTC'
  }
}

/** Local wall-clock ISO 8601 (no offset) — the reference "now" for the AI. */
function localNowIso(): string {
  const d = new Date()
  const pad = (n: number): string => String(n).padStart(2, '0')
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
  )
}

async function findMeetingTimes(): Promise<void> {
  status.value = null
  meetingProposals.value = null
  const r = await run('meeting', () =>
    call((c) =>
      c.extractMeetingTimes({
        subject: item.value.subject,
        body: item.value.bodyText,
        from: item.value.from,
        nowIso: localNowIso(),
        timezone: userTimezone(),
      }),
    ),
  )
  if (r) meetingProposals.value = r
}

function appointmentSupported(): boolean {
  const mb = (typeof Office !== 'undefined' ? Office.context?.mailbox : undefined) as
    | { displayNewAppointmentForm?: unknown }
    | undefined
  return !!mb && typeof mb.displayNewAppointmentForm === 'function'
}

function formatSlot(p: MeetingProposal): string {
  const start = new Date(p.startIso)
  if (Number.isNaN(start.getTime())) return p.title
  const when = start.toLocaleString(undefined, {
    weekday: 'short',
    day: 'numeric',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit',
  })
  return `${p.title} — ${when}`
}

function addToCalendar(p: MeetingProposal): void {
  error.value = null
  if (!appointmentSupported()) {
    error.value = t('read.meeting.unsupported')
    return
  }
  const start = new Date(p.startIso)
  const end = new Date(p.endIso)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    error.value = t('read.meeting.badDate')
    return
  }
  try {
    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: senderEmail.value ? [senderEmail.value] : [],
      start,
      end,
      subject: p.title,
      location: p.location,
      body: t('read.meeting.bodyFrom', { subject: item.value.subject }),
    })
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  }
}
</script>

<template>
  <section class="read">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('home')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ item.subject || '—' }}</h2>
      <p v-if="item.from" class="syn-view-subtitle">{{ t('read.from') }}: {{ item.from }}</p>
    </header>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('read.title') }}</h3>
      <div class="syn-stack">
        <ActionButton :loading="active === 'summarise'" @click="summarise">
          {{ t('read.actions.summarise') }}
        </ActionButton>

        <div class="syn-row">
          <ActionButton :loading="active === 'translate'" :block="false" @click="translate">
            {{ t('read.actions.translate') }}
          </ActionButton>
          <LanguagePicker v-model="targetLang" />
        </div>

        <div class="syn-row">
          <ActionButton :loading="active === 'reply'" :block="false" @click="draftReply">
            {{ t('read.actions.draftReply') }}
          </ActionButton>
          <TonePicker v-model="tone" />
        </div>

        <ActionButton :loading="active === 'classify'" @click="classify">
          {{ t('read.actions.classify') }}
        </ActionButton>

        <ActionButton :loading="active === 'save'" @click="openSaveDialog">
          {{ t('read.actions.saveToRag') }}
        </ActionButton>

        <ActionButton :loading="active === 'meeting'" @click="findMeetingTimes">
          {{ t('read.actions.findMeetingTimes') }}
        </ActionButton>
      </div>
    </div>

    <div v-if="meetingProposals" class="syn-card">
      <h3 class="syn-card-title">{{ t('read.meeting.title') }}</h3>
      <ul v-if="meetingProposals.length" class="read__slots">
        <li v-for="(p, i) in meetingProposals" :key="i">
          <button type="button" class="read__slot" @click="addToCalendar(p)">
            <span>{{ formatSlot(p) }}</span>
            <span v-if="p.location" class="syn-card-sub">{{ p.location }}</span>
          </button>
        </li>
      </ul>
      <p v-else class="syn-muted">{{ t('read.meeting.none') }}</p>
    </div>

    <pre v-if="result" class="read__result">{{ result }}</pre>
    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />

    <div v-if="senderEmail" class="syn-card">
      <h3 class="syn-card-title">{{ t('read.senderTitle') }}</h3>
      <p class="syn-card-sub">{{ senderEmail }}</p>
      <div class="syn-stack">
        <ActionButton @click="openContactKb(senderEmail)">{{ t('read.contactKb') }}</ActionButton>
        <ActionButton :loading="active === 'senderHistory'" @click="loadSenderHistory">
          {{ t('read.actions.moreFromSender') }}
        </ActionButton>
        <ActionButton variant="danger" @click="openBlockDialog">
          {{ t('read.actions.blockSender') }}
        </ActionButton>
      </div>
      <SenderHistoryList
        v-if="senderHistory"
        :history="senderHistory"
        @summarise="summariseSenderHistory"
      />
    </div>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('read.askTitle') }}</h3>
      <p v-if="imageCount > 0" class="syn-card-sub read__image-hint">
        {{ t('read.askImageHint', { n: imageCount }) }}
      </p>
      <div v-for="(turn, i) in askHistory" :key="i" class="read__turn">
        <p><strong>You:</strong> {{ turn.q }}</p>
        <p><strong>Synaplan:</strong> {{ turn.a }}</p>
      </div>
      <div class="syn-row">
        <input
          v-model="question"
          type="text"
          :placeholder="t('read.askPlaceholder')"
          class="read__ask-input"
          @keyup.enter="ask"
        />
        <ActionButton :block="false" :loading="active === 'ask'" @click="ask">
          {{ t('read.actions.ask') }}
        </ActionButton>
      </div>
    </div>

    <SaveToRagDialog
      v-if="showSaveDialog"
      :contact-email="senderEmail"
      :last-used-group-id="getLastRagGroupId()"
      @cancel="showSaveDialog = false"
      @confirm="handleSaveConfirm"
    />

    <BlockSenderDialog
      v-if="showBlockDialog"
      :sender-email="senderEmail"
      @cancel="showBlockDialog = false"
      @confirm="handleBlockConfirm"
    />
  </section>
</template>

<style scoped>
.read {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.read__result {
  background: var(--syn-surface);
  padding: var(--syn-space-3);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  white-space: pre-wrap;
  margin: 0;
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
}
.read__turn {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
}
.read__ask-input {
  flex: 1;
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
}
.read__slots {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.read__slot {
  width: 100%;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: 2px;
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
  cursor: pointer;
}
</style>
