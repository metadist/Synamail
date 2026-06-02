<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import TonePicker from '@/taskpane/components/TonePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { canDisplayNewMessage, displayNewMessage } from '@/taskpane/composables/useOutlookCompose'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

// Empty by design — the user types how they want to answer the sender.
const prompt = ref('')
const targetLang = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('en')
const tone = ref<'formal' | 'concise' | 'friendly'>('concise')
const busy = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

// A read item means there's a sender to reply to; otherwise we compose anew.
const replying = computed(() => item.value.mode === 'read')
const sender = computed(() => item.value.from ?? '')

async function draft(): Promise<void> {
  busy.value = true
  error.value = null
  status.value = null
  try {
    const lang = targetLang.value === 'auto' ? 'en' : targetLang.value
    if (replying.value) {
      // Fold the user's instruction (if any) into the email block so the AI
      // tailors the reply; an empty prompt just yields a standard reply.
      const instruction = prompt.value.trim()
      const body = instruction
        ? `${item.value.bodyText}\n\n[Reply instructions]\n${instruction}`
        : item.value.bodyText
      const r = await call((c) =>
        c.draftReply({ subject: item.value.subject, body, tone: tone.value, language: lang }),
      )
      if (r) {
        Office.context.mailbox.item?.displayReplyForm({ htmlBody: r.htmlBody })
        status.value = t('composeAnswer.replyOpened')
      }
    } else {
      const description = prompt.value.trim()
      if (!description) {
        error.value = t('composeAnswer.needPrompt')
        return
      }
      if (!canDisplayNewMessage()) {
        error.value = t('home.newMail.unsupportedHost')
        return
      }
      const r = await call((c) => c.composeNew({ description, language: lang }))
      if (r) {
        await displayNewMessage({ subject: r.subject, htmlBody: r.htmlBody })
        status.value = t('home.newMail.opened')
      }
    }
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    busy.value = false
  }
}
</script>

<template>
  <div class="ca">
    <p class="syn-card-sub">
      {{ replying ? t('composeAnswer.replyHint', { email: sender }) : t('composeAnswer.newHint') }}
    </p>

    <textarea
      v-model="prompt"
      rows="3"
      :placeholder="
        replying ? t('composeAnswer.replyPlaceholder') : t('composeAnswer.newPlaceholder')
      "
    />

    <div class="syn-row">
      <ActionButton variant="primary" :block="false" :loading="busy" @click="draft">
        {{ replying ? t('composeAnswer.draftReply') : t('composeAnswer.draftNew') }}
      </ActionButton>
      <TonePicker v-if="replying" v-model="tone" />
      <LanguagePicker v-model="targetLang" />
    </div>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </div>
</template>

<style scoped>
.ca {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.ca textarea {
  width: 100%;
  box-sizing: border-box;
}
</style>
