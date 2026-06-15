<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import AccordionItem from '@/taskpane/components/AccordionItem.vue'
import ChatThread from '@/taskpane/components/ChatThread.vue'
import type { ChatMessage } from '@/taskpane/components/ChatThread.vue'
import ContactProfilePanel from '@/taskpane/components/ContactProfilePanel.vue'
import EmailActionsPanel from '@/taskpane/components/EmailActionsPanel.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useContactCounterpart } from '@/taskpane/composables/useContactCounterpart'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import {
  clearChatIdForConversation,
  getChatIdForConversation,
  setChatIdForConversation,
} from '@/taskpane/composables/useRoamingSettings'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

const HOME_CONVERSATION = 'home'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()
const { contactEmail } = useContactCounterpart(item)

const messages = ref<ChatMessage[]>([])
const sending = ref(false)
const error = ref<string | null>(null)

const emailOpen = computed(() => item.value.mode === 'read')

function truncate(text: string, max = 48): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

const emailActionsSubtitle = computed(() =>
  emailOpen.value
    ? `${t('read.subjectLabel')}: ${truncate(item.value.subject || t('home.emailTitle'))}`
    : t('read.noEmail'),
)

const profilingSubtitle = computed(() =>
  contactEmail.value ? truncate(contactEmail.value) : t('read.noEmail'),
)

async function send(text: string): Promise<void> {
  messages.value.push({ role: 'user', text })
  // Add the AI bubble up front and stream tokens into it as they arrive.
  const aiIdx = messages.value.push({ role: 'ai', text: '' }) - 1
  sending.value = true
  error.value = null
  try {
    const chatId = getChatIdForConversation(HOME_CONVERSATION)
    const r = await call((c) =>
      c.chat({ conversationId: HOME_CONVERSATION, question: text, chatId }, (textSoFar) => {
        messages.value[aiIdx].text = textSoFar
      }),
    )
    if (r) {
      messages.value[aiIdx].text = r.answer
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
    <!-- 1. Chat — always open, above the flaps. -->
    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.chat') }}</h2>
      <ChatThread
        :messages="messages"
        :loading="sending"
        :initial-draft="t('home.chat.sample')"
        @send="send"
        @reset="resetChat"
      />
    </div>

    <Toast v-if="error" kind="error" :message="error" />

    <!-- 2. Email actions for the active email. -->
    <AccordionItem
      :title="t('home.sections.emailActions')"
      :subtitle="emailActionsSubtitle"
      :strong-subtitle="emailOpen"
    >
      <EmailActionsPanel />
    </AccordionItem>

    <!-- 3. Contact profiling for the email's counterpart. -->
    <AccordionItem
      :title="t('home.sections.profiling')"
      :subtitle="profilingSubtitle"
      :strong-subtitle="!!contactEmail"
    >
      <ContactProfilePanel />
    </AccordionItem>
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
