<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { renderMarkdown } from '@shared/markdown'

export interface ChatMessage {
  role: 'user' | 'ai'
  text: string
}

interface Props {
  messages: ChatMessage[]
  loading?: boolean
  placeholder?: string
  emptyHint?: string
  /** Seed the composer so Send is active and the user sees a concrete example. */
  initialDraft?: string
}

const props = withDefaults(defineProps<Props>(), {
  loading: false,
  placeholder: '',
  emptyHint: '',
  initialDraft: '',
})

const emit = defineEmits<{
  (e: 'send', text: string, fileIds?: number[]): void
  (e: 'reset'): void
}>()

const { t, locale } = useI18n()
const { call } = useSynaplanClient()
const draft = ref(props.initialDraft)
const root = ref<HTMLElement | null>(null)

// True when the last message is an AI bubble still waiting for its first token
// (the caller streams into it). The thinking indicator then lives in that
// bubble instead of a separate one, so they never show at the same time.
const pendingAi = computed(() => {
  const last = props.messages[props.messages.length - 1]
  return !!last && last.role === 'ai' && !last.text
})

// The composer is seeded with an example so Send is active and users see they
// can type here. The first time they focus the still-untouched sample, clear it
// so they don't have to delete it by hand. Once they've focused or edited it,
// we never auto-clear again — so it's never annoying on later interactions.
const sampleActive = ref(Boolean(props.initialDraft))

function onFocus(): void {
  if (sampleActive.value && draft.value === props.initialDraft) {
    draft.value = ''
  }
  sampleActive.value = false
}

function onInput(): void {
  sampleActive.value = false
}

// ---------------------------------------------------------------------------
// Attachments — the "+" button uploads files to Synaplan and passes their ids
// with the next question (the AI gets each file's extracted text / vision).
// ---------------------------------------------------------------------------

interface Attachment {
  name: string
  fileId: number
}

const fileInput = ref<HTMLInputElement | null>(null)
const attachments = ref<Attachment[]>([])
const uploading = ref(false)
const uploadError = ref<string | null>(null)

function pickFiles(): void {
  fileInput.value?.click()
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const result = String(reader.result ?? '')
      const comma = result.indexOf(',')
      resolve(comma >= 0 ? result.slice(comma + 1) : result)
    }
    reader.onerror = () => reject(reader.error ?? new Error('file read failed'))
    reader.readAsDataURL(file)
  })
}

async function onFilesSelected(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement
  const files = Array.from(input.files ?? [])
  input.value = ''
  if (files.length === 0) return
  uploadError.value = null
  uploading.value = true
  try {
    for (const file of files) {
      const contentBase64 = await fileToBase64(file)
      const r = await call((c) =>
        c.fileUpload({
          filename: file.name,
          contentBase64,
          mimeType: file.type || 'application/octet-stream',
          processLevel: 'extract',
        }),
      )
      if (r?.fileId) attachments.value.push({ name: file.name, fileId: r.fileId })
    }
  } catch (err) {
    uploadError.value = errorMessage(err)
  } finally {
    uploading.value = false
  }
}

function removeAttachment(index: number): void {
  attachments.value.splice(index, 1)
}

// ---------------------------------------------------------------------------
// Dictation — the microphone uses the browser's Web Speech API to transcribe
// speech into the composer. Hidden entirely when the API is unavailable.
// ---------------------------------------------------------------------------

interface SpeechAlternative {
  transcript: string
}
interface SpeechResult {
  readonly length: number
  [index: number]: SpeechAlternative
}
interface SpeechResultList {
  readonly length: number
  [index: number]: SpeechResult
}
interface SpeechRecognitionEventLike {
  results: SpeechResultList
}
interface SpeechRecognitionLike {
  lang: string
  interimResults: boolean
  continuous: boolean
  onresult: ((event: SpeechRecognitionEventLike) => void) | null
  onend: (() => void) | null
  onerror: (() => void) | null
  start(): void
  stop(): void
  abort(): void
}
type SpeechRecognitionCtor = new () => SpeechRecognitionLike

const micSupported = ref(false)
const recording = ref(false)
let recognition: SpeechRecognitionLike | null = null

function speechCtor(): SpeechRecognitionCtor | undefined {
  const w = window as unknown as {
    SpeechRecognition?: SpeechRecognitionCtor
    webkitSpeechRecognition?: SpeechRecognitionCtor
  }
  return w.SpeechRecognition ?? w.webkitSpeechRecognition
}

function toggleMic(): void {
  if (recording.value) {
    recognition?.stop()
    return
  }
  const Ctor = speechCtor()
  if (!Ctor) return
  const rec = new Ctor()
  rec.lang = locale.value
  rec.interimResults = true
  rec.continuous = false
  const startingFromSample = sampleActive.value && draft.value === props.initialDraft
  const base = startingFromSample || !draft.value ? '' : `${draft.value} `
  rec.onresult = (event) => {
    let transcript = ''
    for (let i = 0; i < event.results.length; i++) {
      transcript += event.results[i][0].transcript
    }
    draft.value = base + transcript
    sampleActive.value = false
  }
  rec.onend = () => {
    recording.value = false
  }
  rec.onerror = () => {
    recording.value = false
  }
  recognition = rec
  recording.value = true
  rec.start()
}

function submit(): void {
  const text = draft.value.trim()
  if (!text || props.loading) return
  const fileIds = attachments.value.map((a) => a.fileId)
  emit('send', text, fileIds.length ? fileIds : undefined)
  draft.value = ''
  attachments.value = []
  sampleActive.value = false
}

function scrollToTop(): void {
  root.value?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

// Keep the newest reply — and the sticky composer — in view as the thread grows.
function scrollToLatest(): void {
  try {
    globalThis.scrollTo?.({ top: document.documentElement.scrollHeight, behavior: 'smooth' })
  } catch {
    /* jsdom / non-browser env has no real scrolling */
  }
}

watch(
  () => props.messages.length,
  async () => {
    await nextTick()
    scrollToLatest()
  },
)

onMounted(() => {
  micSupported.value = !!speechCtor()
})

onBeforeUnmount(() => {
  try {
    recognition?.abort()
  } catch {
    /* recognition may already be stopped */
  }
})
</script>

<template>
  <div ref="root" class="chat">
    <div v-if="messages.length" class="chat__toolbar">
      <button type="button" class="chat__tool" @click="scrollToTop">
        ↑ {{ t('home.chat.scrollTop') }}
      </button>
      <button type="button" class="chat__tool" @click="$emit('reset')">
        {{ t('home.chat.reset') }}
      </button>
    </div>

    <!-- Results render above the composer (newest at the bottom of the list). -->
    <div class="chat__scroll">
      <p v-if="messages.length === 0 && emptyHint" class="syn-muted chat__empty">
        {{ emptyHint }}
      </p>
      <div v-for="(m, i) in messages" :key="i" :class="['chat__bubble', `chat__bubble--${m.role}`]">
        <template v-if="m.role === 'ai'">
          <!-- While a streamed bubble is still empty, it shows the thinking
               indicator; once tokens arrive it renders Markdown. renderMarkdown
               sanitises the text, so v-html is XSS-safe here. -->
          <span v-if="!m.text && loading && i === messages.length - 1" class="chat__loading">
            <span class="chat__dots" aria-hidden="true">…</span>
            {{ t('home.chat.thinking') }}
          </span>
          <!-- eslint-disable-next-line vue/no-v-html -->
          <div v-else class="syn-md" v-html="renderMarkdown(m.text)" />
        </template>
        <template v-else>{{ m.text }}</template>
      </div>
      <div v-if="loading && !pendingAi" class="chat__bubble chat__bubble--ai chat__bubble--loading">
        <span class="chat__dots" aria-hidden="true">…</span>
        {{ t('home.chat.thinking') }}
      </div>
    </div>

    <!-- Composer mimics the Synaplan chat box: a shell with the input, a "+"
         (attach) on the left and a microphone + Send on the right. -->
    <div class="chat__composer" :class="{ 'chat__composer--bordered': messages.length }">
      <div v-if="attachments.length || uploading" class="chat__attachments">
        <span v-if="uploading" class="syn-muted">{{ t('home.chat.uploading') }}</span>
        <span v-else class="syn-muted">
          {{ t('home.chat.attachments', { n: attachments.length }) }}
        </span>
        <button
          v-for="(a, i) in attachments"
          :key="i"
          type="button"
          class="chat__chip"
          @click="removeAttachment(i)"
        >
          {{ a.name }} <span aria-hidden="true">✕</span>
        </button>
      </div>

      <div class="chat__shell">
        <textarea
          v-model="draft"
          rows="2"
          class="chat__input"
          :placeholder="placeholder || t('home.chat.placeholder')"
          @focus="onFocus"
          @input="onInput"
          @keydown.enter.exact.prevent="submit"
        />
        <div class="chat__controls">
          <button
            type="button"
            class="chat__icon"
            :aria-label="t('home.chat.attach')"
            :title="t('home.chat.attach')"
            @click="pickFiles"
          >
            <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true">
              <path
                d="M10 4v12M4 10h12"
                stroke="currentColor"
                stroke-width="2"
                stroke-linecap="round"
                fill="none"
              />
            </svg>
          </button>
          <input
            ref="fileInput"
            type="file"
            multiple
            class="chat__file"
            @change="onFilesSelected"
          />

          <div class="chat__controls-right">
            <button
              v-if="micSupported"
              type="button"
              class="chat__icon"
              :class="{ 'chat__icon--rec': recording }"
              :aria-label="recording ? t('home.chat.stopDictation') : t('home.chat.dictate')"
              :title="recording ? t('home.chat.stopDictation') : t('home.chat.dictate')"
              @click="toggleMic"
            >
              <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true">
                <path
                  d="M10 2a2.5 2.5 0 0 0-2.5 2.5v5a2.5 2.5 0 0 0 5 0v-5A2.5 2.5 0 0 0 10 2z"
                  fill="currentColor"
                />
                <path
                  d="M5 9.5a5 5 0 0 0 10 0M10 14.5V18M7 18h6"
                  stroke="currentColor"
                  stroke-width="1.5"
                  stroke-linecap="round"
                  fill="none"
                />
              </svg>
            </button>
            <ActionButton
              variant="primary"
              :block="false"
              :loading="loading"
              :disabled="!draft.trim()"
              @click="submit"
            >
              {{ t('home.chat.send') }}
            </ActionButton>
          </div>
        </div>
      </div>

      <Toast v-if="uploadError" kind="error" :message="uploadError" />
    </div>
  </div>
</template>

<style scoped>
.chat {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.chat__toolbar {
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  justify-content: flex-end;
  gap: var(--syn-space-3);
  padding-bottom: var(--syn-space-1);
  background: var(--syn-surface);
  border-bottom: 1px solid var(--syn-border);
}
.chat__tool {
  background: none;
  border: 0;
  padding: 0;
  color: var(--syn-brand-600);
  font-size: var(--syn-font-size-sm);
  font-weight: 500;
}
.chat__tool:hover {
  text-decoration: underline;
}
.chat__scroll {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  padding: var(--syn-space-1);
}
.chat__empty {
  margin: 0;
  text-align: left;
  font-size: var(--syn-font-size-sm);
}
.chat__bubble {
  padding: var(--syn-space-2) var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  font-size: var(--syn-font-size-sm);
  white-space: pre-wrap;
  word-break: break-word;
  max-width: 92%;
}
.chat__bubble--user {
  align-self: flex-end;
  background: var(--syn-brand-600);
  color: white;
}
.chat__bubble--ai {
  align-self: flex-start;
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
}
.chat__bubble--loading {
  opacity: 0.8;
}
/* Markdown formatting comes from the shared `.syn-md` rules in app.css. */
.chat__dots {
  margin-right: var(--syn-space-1);
}
.chat__composer {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
/* Only divide the composer from the thread once messages exist — on the empty
   initial state the line under the non-foldable headline looks out of place. */
.chat__composer--bordered {
  padding-top: var(--syn-space-2);
  border-top: 1px solid var(--syn-border);
}
.chat__attachments {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--syn-space-2);
}
.chat__chip {
  display: inline-flex;
  align-items: center;
  gap: var(--syn-space-1);
  max-width: 100%;
  padding: 2px var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: 999px;
  background: var(--syn-bg);
  color: var(--syn-text);
  font-size: var(--syn-font-size-xs);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
/* Chat input shell — a rounded card holding the textarea and its controls,
   echoing the Synaplan composer. */
.chat__shell {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
  padding: var(--syn-space-2);
  background: var(--syn-bg);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-lg);
}
.chat__shell:focus-within {
  border-color: var(--syn-brand-500);
  outline: 1px solid var(--syn-brand-500);
}
.chat__input {
  width: 100%;
  border: 0;
  background: transparent;
  padding: var(--syn-space-1);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
  resize: none;
}
.chat__input:focus-visible {
  outline: none;
  border: 0;
}
.chat__controls {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--syn-space-2);
}
.chat__controls-right {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
}
.chat__file {
  display: none;
}
.chat__icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 36px;
  height: 36px;
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  background: var(--syn-surface);
  color: var(--syn-text);
}
.chat__icon:hover:not(:disabled) {
  border-color: var(--syn-brand-500);
  color: var(--syn-brand-600);
}
.chat__icon--rec {
  background: var(--syn-danger);
  border-color: var(--syn-danger);
  color: white;
}
</style>
