<script setup lang="ts">
import { ref } from 'vue'
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
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

const HOME_CONVERSATION = 'home'

const { t } = useI18n()
const { call } = useSynaplanClient()

const messages = ref<ChatMessage[]>([])
const sending = ref(false)
const error = ref<string | null>(null)

async function send(text: string, fileIds?: number[]): Promise<void> {
  messages.value.push({ role: 'user', text })
  // Add the AI bubble up front and stream tokens into it as they arrive.
  const aiIdx = messages.value.push({ role: 'ai', text: '' }) - 1
  sending.value = true
  error.value = null
  try {
    const chatId = getChatIdForConversation(HOME_CONVERSATION)
    const r = await call((c) =>
      c.chat(
        { conversationId: HOME_CONVERSATION, question: text, chatId, fileIds },
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
          await setChatIdForConversation(HOME_CONVERSATION, r.chatId)
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
  messages.value = []
  error.value = null
  try {
    await clearChatIdForConversation(HOME_CONVERSATION)
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

    <!-- (d) Ask Synaplan — general chat with results above the composer. -->
    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.chat') }}</h2>
      <ChatThread :messages="messages" :loading="sending" @send="send" @reset="resetChat" />
      <Toast v-if="error" kind="error" :message="error" />
    </div>

    <!-- Profiling is temporarily disabled and will return in a later iteration.
    <AccordionItem :title="t('home.sections.profiling')">
      <ContactProfilePanel />
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
