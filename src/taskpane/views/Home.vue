<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import ChatThread from '@/taskpane/components/ChatThread.vue'
import type { ChatMessage } from '@/taskpane/components/ChatThread.vue'
import MailSearchVectorizeDialog from '@/taskpane/components/MailSearchVectorizeDialog.vue'
import NewMailDialog from '@/taskpane/components/NewMailDialog.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import {
  clearChatIdForConversation,
  getChatIdForConversation,
  getLastRagGroupId,
  setChatIdForConversation,
} from '@/taskpane/composables/useRoamingSettings'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { go } from '@/taskpane/router'

const HOME_CONVERSATION = 'home'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

const messages = ref<ChatMessage[]>([])
const sending = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const showSearch = ref(false)
const showNewMail = ref(false)

const emailLoaded = computed(() => item.value.mode === 'read' || item.value.mode === 'compose')

async function send(text: string): Promise<void> {
  messages.value.push({ role: 'user', text })
  sending.value = true
  error.value = null
  try {
    const chatId = getChatIdForConversation(HOME_CONVERSATION)
    const r = await call((c) =>
      c.chat({ conversationId: HOME_CONVERSATION, question: text, chatId }),
    )
    if (r) {
      messages.value.push({ role: 'ai', text: r.answer })
      if (!chatId && r.chatId) {
        try {
          await setChatIdForConversation(HOME_CONVERSATION, r.chatId)
        } catch {
          // Roaming write may fail offline/in tests; the in-memory thread
          // still works for the rest of this session.
        }
      }
    }
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    sending.value = false
  }
}

function onVectorizeDone(group: string): void {
  showSearch.value = false
  status.value = t('home.search.done', { group })
}

function onMailOpened(): void {
  showNewMail.value = false
  status.value = t('home.newMail.opened')
}

async function resetChat(): Promise<void> {
  messages.value = []
  error.value = null
  // Start a fresh Synaplan chat on the next message.
  try {
    await clearChatIdForConversation(HOME_CONVERSATION)
  } catch {
    // Roaming write can fail offline/in tests; the in-memory reset still holds.
  }
}
</script>

<template>
  <section class="home">
    <div v-if="emailLoaded" class="syn-card">
      <h2 class="syn-card-title">{{ t('home.emailTitle') }}</h2>
      <p class="syn-card-sub">{{ item.subject || '—' }}</p>
      <ActionButton @click="go(item.mode === 'compose' ? 'compose' : 'read')">
        {{ t('home.emailActions') }}
      </ActionButton>
    </div>

    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.search') }}</h2>
      <p class="syn-card-sub">{{ t('home.commands.searchSub') }}</p>
      <ActionButton data-testid="cmd-search" @click="showSearch = true">
        {{ t('home.commands.searchAction') }}
      </ActionButton>
    </div>

    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.newMail') }}</h2>
      <p class="syn-card-sub">{{ t('home.commands.newMailSub') }}</p>
      <ActionButton data-testid="cmd-newmail" @click="showNewMail = true">
        {{ t('home.commands.newMailAction') }}
      </ActionButton>
    </div>

    <Toast v-if="status" kind="success" :message="status" />

    <div class="syn-card">
      <h2 class="syn-card-title">{{ t('home.commands.chat') }}</h2>
      <p class="syn-card-sub">{{ t('home.chat.emptyHint') }}</p>
      <ChatThread
        :messages="messages"
        :loading="sending"
        :initial-draft="t('home.chat.sample')"
        @send="send"
        @reset="resetChat"
      />
    </div>

    <Toast v-if="error" kind="error" :message="error" />

    <MailSearchVectorizeDialog
      v-if="showSearch"
      :last-used-group-id="getLastRagGroupId()"
      @cancel="showSearch = false"
      @done="onVectorizeDone"
    />
    <NewMailDialog v-if="showNewMail" @cancel="showNewMail = false" @opened="onMailOpened" />
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
