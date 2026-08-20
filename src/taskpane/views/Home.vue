<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ChatThread from '@/taskpane/components/ChatThread.vue'
import type { ChatMessage } from '@/taskpane/components/ChatThread.vue'
import EmailWritingBox from '@/taskpane/components/EmailWritingBox.vue'
import KnowledgeBaseBox from '@/taskpane/components/KnowledgeBaseBox.vue'
import SummarizeBox from '@/taskpane/components/SummarizeBox.vue'
import Toast from '@/taskpane/components/Toast.vue'
import {
  clearChatIdForConversation,
  getChatIdForConversation,
  setChatIdForConversation,
} from '@/taskpane/composables/useRoamingSettings'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage, isApiError } from '@shared/synaplan-client'

const { t } = useI18n()
const { call } = useSynaplanClient()
const { item } = useOutlookItem()

/**
 * Per-mail chat key (FEATURES.md §1.5). Prefer Outlook's conversationId so
 * close/reopen on the same mail restores the same Synaplan thread; fall back
 * to subject, then a stable home key when nothing is open.
 */
const conversationKey = computed(() => {
  if (item.value.conversationId) return item.value.conversationId
  const subject = item.value.subject.trim()
  if (subject) return `synamail:${subject}`
  return 'home'
})

const messages = ref<ChatMessage[]>([])
const sending = ref(false)
const restoring = ref(false)
const error = ref<string | null>(null)

// Monotonic token identifying the LATEST restore. Only the latest call may
// drive `restoring`/`messages` — a superseded in-flight restore must not
// touch them. (Keying this off `conversationKey` equality deadlocked once:
// on load the key flips from 'home' to the mail's conversationId while the
// 'home' restore is still in flight; the new key had no saved chat and
// returned early, and the stale finally refused to clear `restoring` because
// the key had changed — leaving the chat spinner stuck until a remount.)
let restoreSeq = 0

async function restoreThread(key: string): Promise<void> {
  const seq = ++restoreSeq
  const chatId = getChatIdForConversation(key)
  if (!chatId) {
    messages.value = []
    // We're the latest restore: clear any flag a superseded call left set.
    restoring.value = false
    return
  }
  restoring.value = true
  error.value = null
  try {
    const history = await call((c) => c.getChatMessages(chatId))
    // Only seed when no newer restore has started (ItemChanged can race).
    if (seq !== restoreSeq) return
    messages.value = (history ?? []).map((m) => ({ role: m.role, text: m.text }))
  } catch (err) {
    if (seq !== restoreSeq) return
    messages.value = []
    // Only drop the roaming id when the server says the chat is gone.
    if (isApiError(err) && err.status === 404) {
      try {
        await clearChatIdForConversation(key)
      } catch {
        /* roaming write may fail offline/in tests */
      }
    } else {
      error.value = errorMessage(err)
    }
  } finally {
    if (seq === restoreSeq) restoring.value = false
  }
}

watch(
  conversationKey,
  (key) => {
    void restoreThread(key)
  },
  { immediate: true },
)

async function send(text: string, fileIds?: number[]): Promise<void> {
  messages.value.push({ role: 'user', text })
  // Add the AI bubble up front and stream tokens into it as they arrive.
  const aiIdx = messages.value.push({ role: 'ai', text: '' }) - 1
  sending.value = true
  error.value = null
  try {
    const key = conversationKey.value
    const chatId = getChatIdForConversation(key)
    // Ground on the open mail body when readable (#55); keep general chat
    // when compose is empty or the body could not be read.
    const body = item.value.bodyText.trim()
    const emailContext = body.length > 0 ? body : undefined
    const mailSubject = item.value.subject.trim() || undefined
    const r = await call((c) =>
      c.chat(
        {
          conversationId: key,
          question: text,
          chatId,
          fileIds,
          emailContext,
          mailSubject,
        },
        (textSoFar) => {
          messages.value[aiIdx].text = textSoFar
        },
      ),
    )
    if (r) {
      messages.value[aiIdx].text = r.answer
      if (r.media && r.media.length) messages.value[aiIdx].media = r.media
      if (!chatId && r.chatId) {
        try {
          await setChatIdForConversation(key, r.chatId)
        } catch {
          // Roaming write may fail offline/in tests; the in-memory thread still works.
        }
      }
    } else {
      // 401/cleared client — drop the empty AI bubble.
      messages.value.splice(aiIdx, 1)
    }
  } catch (err) {
    messages.value.splice(aiIdx, 1)
    error.value = errorMessage(err)
  } finally {
    sending.value = false
  }
}

async function resetChat(): Promise<void> {
  const key = conversationKey.value
  messages.value = []
  error.value = null
  try {
    await clearChatIdForConversation(key)
  } catch {
    // Roaming write can fail offline/in tests; the in-memory reset still holds.
  }
}
</script>

<template>
  <section class="home">
    <!-- One box per function. Scope is intentionally narrow for now; more
         tools (and Profiling) return over time. -->

    <!-- (a) Write an email or reply. -->
    <EmailWritingBox />

    <!-- (b) Summarize the open email. -->
    <SummarizeBox />

    <!-- (c) Save the open email to the knowledge base. -->
    <KnowledgeBaseBox />

    <!-- (d) Ask Synaplan — grounded on the open mail when present (#55). -->
    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.chat') }}</h2>
      <ChatThread
        :messages="messages"
        :loading="sending || restoring"
        @send="send"
        @reset="resetChat"
      />
      <Toast v-if="error" kind="error" :message="error" />
    </div>

    <!-- Profiling is temporarily disabled and will return in a later iteration.
    <AccordionItem :title="t('home.sections.profiling')">
      <ContactProfileView />
    </AccordionItem>
    -->
  </section>
</template>

<style scoped>
.home {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
</style>
