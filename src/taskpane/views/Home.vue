<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ChatThread from '@/taskpane/components/ChatThread.vue'
import type { ChatMessage } from '@/taskpane/components/ChatThread.vue'
import MailSearchVectorizeDialog from '@/taskpane/components/MailSearchVectorizeDialog.vue'
import NewMailDialog from '@/taskpane/components/NewMailDialog.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import {
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
</script>

<template>
  <section class="home">
    <p class="syn-muted home__intro">{{ t('home.intro') }}</p>

    <div v-if="emailLoaded" class="home__context">
      <span class="home__context-text">
        {{ t('home.emailLoaded', { subject: item.subject || '—' }) }}
      </span>
      <button
        type="button"
        class="home__context-link"
        @click="go(item.mode === 'compose' ? 'compose' : 'read')"
      >
        {{ t('home.emailActions') }} →
      </button>
    </div>

    <div class="home__commands">
      <button type="button" class="home__cmd" @click="showSearch = true">
        <span class="home__cmd-title">{{ t('home.commands.search') }}</span>
        <span class="home__cmd-sub syn-muted">{{ t('home.commands.searchSub') }}</span>
      </button>
      <button type="button" class="home__cmd" @click="showNewMail = true">
        <span class="home__cmd-title">{{ t('home.commands.newMail') }}</span>
        <span class="home__cmd-sub syn-muted">{{ t('home.commands.newMailSub') }}</span>
      </button>
    </div>

    <Toast v-if="status" kind="success" :message="status" />

    <h2 class="home__chat-title">{{ t('home.commands.chat') }}</h2>
    <ChatThread
      :messages="messages"
      :loading="sending"
      :empty-hint="t('home.chat.emptyHint')"
      @send="send"
    />

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
.home__intro {
  margin: 0;
  font-size: var(--syn-font-size-sm);
}
.home__context {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: var(--syn-space-2);
  padding: var(--syn-space-2) var(--syn-space-3);
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  font-size: var(--syn-font-size-sm);
}
.home__context-text {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.home__context-link {
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  text-decoration: underline;
  cursor: pointer;
  font-size: var(--syn-font-size-sm);
  white-space: nowrap;
}
.home__commands {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.home__cmd {
  display: flex;
  flex-direction: column;
  gap: 2px;
  text-align: left;
  padding: var(--syn-space-3);
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  cursor: pointer;
  color: var(--syn-text);
}
.home__cmd:hover {
  border-color: var(--syn-brand-300);
}
.home__cmd-title {
  font-weight: 500;
}
.home__cmd-sub {
  font-size: var(--syn-font-size-sm);
}
.home__chat-title {
  margin: 0;
  font-size: var(--syn-font-size-lg);
}
</style>
