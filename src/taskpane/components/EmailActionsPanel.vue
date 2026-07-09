<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import SaveToRagDialog from '@/taskpane/components/SaveToRagDialog.vue'
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
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { renderMarkdown } from '@shared/markdown'
import type { MeetingProposal } from '@shared/types'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

type ActionKey =
  | 'summarise'
  | 'translate'
  | 'reply'
  | 'classify'
  | 'save'
  | 'ask'
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

const emailOpen = computed(() => item.value.mode === 'read')
const senderEmail = computed(() => item.value.from ?? '')
const conversationKey = computed(
  () => item.value.conversationId ?? `synamail:${item.value.subject}`,
)
const imageCount = computed(() => countImageAttachments(item.value))
const imageFileIds = ref<number[]>([])

async function run<T>(key: NonNullable<ActionKey>, fn: () => Promise<T | null>): Promise<T | null> {
  active.value = key
  error.value = null
  try {
    return await fn()
  } catch (err) {
    error.value = errorMessage(err)
    return null
  } finally {
    active.value = null
  }
}

async function summarise(): Promise<void> {
  result.value = ''
  const r = await run('summarise', () =>
    call((c) =>
      c.summarise(
        {
          subject: item.value.subject,
          body: item.value.bodyText,
          from: item.value.from,
          to: item.value.to,
        },
        (textSoFar) => {
          result.value = textSoFar
        },
      ),
    ),
  )
  if (r) result.value = r.summary
}

async function translate(): Promise<void> {
  result.value = ''
  const r = await run('translate', () =>
    call((c) =>
      c.translate(
        {
          text: item.value.bodyText,
          targetLanguage: targetLang.value === 'auto' ? 'en' : targetLang.value,
        },
        (textSoFar) => {
          result.value = textSoFar
        },
      ),
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
      error.value = errorMessage(err)
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
    const file = getReadItemAsFile(item.value)
    const upload = (processLevel: 'store' | 'extract' | 'vectorize' | 'full') =>
      call((c) =>
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
          processLevel,
        }),
      )

    // `vectorize`/`full` need an embedding model + reachable vector store, which
    // some workspaces don't have — that makes the whole upload fail server-side.
    // Retry once at `extract` so the email is still saved (just not vectorized)
    // instead of losing it, and tell the user what happened.
    let vectorizeFailed = false
    let r: Awaited<ReturnType<typeof upload>>
    if (payload.processLevel === 'vectorize' || payload.processLevel === 'full') {
      try {
        r = await upload(payload.processLevel)
      } catch {
        vectorizeFailed = true
        r = await upload('extract')
      }
    } else {
      r = await upload(payload.processLevel)
    }

    if (r) {
      await setLastRagGroupId(payload.groupId)
      status.value = vectorizeFailed
        ? t('read.saveDialog.savedNoVectorize', { group: payload.groupId, fileId: r.fileId })
        : t('read.saveDialog.savedTo', { group: payload.groupId, fileId: r.fileId })
    }
    return r
  })
}

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
  // Add the turn up front and stream the answer into it so the reply renders
  // as it arrives instead of after a long wait.
  const idx = askHistory.value.push({ q, a: '' }) - 1
  question.value = ''
  const r = await run('ask', async () => {
    const fileIds = await ensureImageFileIds()
    return call((c) =>
      c.ask(
        {
          conversationId,
          question: q,
          emailContext: item.value.bodyText,
          chatId,
          fileIds: fileIds.length ? fileIds : undefined,
        },
        (textSoFar) => {
          askHistory.value[idx].a = textSoFar
        },
      ),
    )
  })
  if (r) {
    if (!chatId && r.chatId) {
      try {
        await setChatIdForConversation(conversationId, r.chatId)
      } catch {
        // Roaming write can fail in tests / offline; in-memory chatId still works.
      }
    }
    askHistory.value[idx].a = r.answer
  } else {
    // Failed (or 401) — drop the empty turn and restore the question to retry.
    askHistory.value.splice(idx, 1)
    question.value = q
  }
}

function userTimezone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC'
  } catch {
    return 'UTC'
  }
}

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
    error.value = errorMessage(err)
  }
}
</script>

<template>
  <div class="ea">
    <p v-if="!emailOpen" class="syn-muted">{{ t('read.noEmail') }}</p>

    <template v-else>
      <!-- Feedback sits at the top, right where the action buttons are, so the
           result of a click is visible without scrolling to the bottom. -->
      <Toast v-if="status" kind="success" :message="status" />
      <Toast v-if="error" kind="error" :message="error" />

      <!-- One row per action: label on the left, an optional control + a
           uniform Go button on the right. Each row wraps independently — when
           the pane is too narrow to fit the label and the button on one line,
           the control + Go group drops to the next line (still right-aligned)
           instead of overflowing and clipping the button. -->
      <div class="ea__actions">
        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.summarise') }}</span>
          <div class="ea__row-end">
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'summarise'"
              @click="summarise"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>

        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.translate') }}</span>
          <div class="ea__row-end">
            <LanguagePicker v-model="targetLang" />
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'translate'"
              @click="translate"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>

        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.draftReply') }}</span>
          <div class="ea__row-end">
            <TonePicker v-model="tone" />
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'reply'"
              @click="draftReply"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>

        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.classify') }}</span>
          <div class="ea__row-end">
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'classify'"
              @click="classify"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>

        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.saveToRag') }}</span>
          <div class="ea__row-end">
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'save'"
              @click="openSaveDialog"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>

        <div class="ea__row">
          <span class="ea__label">{{ t('read.actions.findMeetingTimes') }}</span>
          <div class="ea__row-end">
            <ActionButton
              class="ea__go"
              :block="false"
              :loading="active === 'meeting'"
              @click="findMeetingTimes"
            >
              {{ t('common.go') }}
            </ActionButton>
          </div>
        </div>
      </div>

      <div v-if="meetingProposals" class="ea__block">
        <h4 class="syn-card-title">{{ t('read.meeting.title') }}</h4>
        <ul v-if="meetingProposals.length" class="ea__slots">
          <li v-for="(p, i) in meetingProposals" :key="i">
            <button type="button" class="ea__slot" @click="addToCalendar(p)">
              <span>{{ formatSlot(p) }}</span>
              <span v-if="p.location" class="syn-card-sub">{{ p.location }}</span>
            </button>
          </li>
        </ul>
        <p v-else class="syn-muted">{{ t('read.meeting.none') }}</p>
      </div>

      <!-- AI output is Markdown/HTML; renderMarkdown sanitises it (strips raw
           tags, escapes the rest), so v-html is safe here. -->
      <!-- eslint-disable-next-line vue/no-v-html -->
      <div v-if="result" class="ea__result syn-md" v-html="renderMarkdown(result)" />

      <div class="ea__block">
        <h4 class="syn-card-title">{{ t('read.askTitle') }}</h4>
        <p v-if="imageCount > 0" class="syn-card-sub">
          {{ t('read.askImageHint', { n: imageCount }) }}
        </p>
        <div v-for="(turn, i) in askHistory" :key="i" class="ea__turn">
          <p class="ea__turn-q">{{ turn.q }}</p>
          <!-- AI answer is Markdown/HTML; renderMarkdown sanitises it. -->
          <!-- eslint-disable-next-line vue/no-v-html -->
          <div class="syn-md" v-html="renderMarkdown(turn.a)" />
        </div>
        <div class="syn-row">
          <input
            v-model="question"
            type="text"
            :placeholder="t('read.askPlaceholder')"
            class="ea__ask-input"
            @keyup.enter="ask"
          />
          <ActionButton :block="false" :loading="active === 'ask'" @click="ask">
            {{ t('read.actions.ask') }}
          </ActionButton>
        </div>
      </div>
    </template>

    <SaveToRagDialog
      v-if="showSaveDialog"
      :contact-email="senderEmail"
      :last-used-group-id="getLastRagGroupId()"
      @cancel="showSaveDialog = false"
      @confirm="handleSaveConfirm"
    />
  </div>
</template>

<style scoped>
.ea {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
/* Action list: one wrapping row per action. The label grows to fill the line;
   the control + Go group is pushed to the right (margin-left:auto) and wraps to
   its own line when the pane is too narrow to fit both — so the Go button is
   never clipped at the edge. The right-alignment keeps every Go button flush
   with the right edge whether or not the row also has a picker. */
.ea__actions {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.ea__row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  column-gap: var(--syn-space-2);
  row-gap: var(--syn-space-1);
}
.ea__label {
  flex: 1 1 8rem;
  min-width: 0;
  font-weight: 500;
  color: var(--syn-text);
}
.ea__label::after {
  content: ':';
}
.ea__row-end {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: var(--syn-space-2);
  margin-left: auto;
}
.ea__row-end :deep(select) {
  min-width: 96px;
}
.ea__go {
  min-width: 56px;
  justify-content: center;
}
.ea__block {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  border-top: 1px solid var(--syn-border);
  padding-top: var(--syn-space-3);
}
.ea__result {
  background: var(--syn-bg);
  padding: var(--syn-space-3);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  font-size: var(--syn-font-size-sm);
}
.ea__turn {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-size: var(--syn-font-size-sm);
}
/* The user's question, set apart from the rendered AI answer below it. */
.ea__turn-q {
  margin: 0;
  font-weight: 600;
  color: var(--syn-muted);
}
.ea__ask-input {
  flex: 1;
}
.ea__slots {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.ea__slot {
  width: 100%;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: 2px;
  background: var(--syn-bg);
  border: 1px solid var(--syn-border);
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
  cursor: pointer;
}
</style>
