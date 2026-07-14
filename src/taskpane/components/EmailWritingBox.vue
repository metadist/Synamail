<script setup lang="ts">
import { computed, nextTick, onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import {
  displayNewMessageWithBody,
  displayReplyWithBody,
  setComposeBody,
  useOutlookItem,
} from '@/taskpane/composables/useOutlookItem'
import { standardLanguage } from '@/taskpane/composables/useLanguagePrefs'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import type { EmailTone } from '@shared/types'

// The three surfaced writing styles (knapp / ausführlich / formell), reusing
// the shared `tone.*` labels so all six locales stay in sync.
const STYLES: EmailTone[] = ['concise', 'detailed', 'formal']

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

const topic = ref('')
const topicInput = ref<HTMLInputElement | null>(null)
const active = ref<EmailTone | null>(null)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

// This is the first box on the taskpane, so give it the cursor as soon as the
// pane opens — the user can start typing the email topic straight away.
onMounted(async () => {
  await nextTick()
  topicInput.value?.focus()
})

// A reply when reading an email, or when the open draft already carries
// text (a reply/forward compose window with the quoted original); otherwise
// a fresh new email.
const isReply = computed(
  () =>
    item.value.mode === 'read' ||
    (item.value.mode === 'compose' && item.value.bodyText.trim().length > 0),
)

async function generate(style: EmailTone): Promise<void> {
  active.value = style
  error.value = null
  status.value = null
  try {
    const r = await call((c) =>
      c.composeDraft({
        intent: topic.value.trim(),
        tone: style,
        language: standardLanguage(),
        // Ground the AI in the message we're replying to (read mode) or the
        // quoted original already in the compose window.
        referenceBody: isReply.value ? item.value.bodyText || undefined : undefined,
      }),
    )
    if (r) await deliver(r.htmlBody)
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    active.value = null
  }
}

// Show the generated draft where it belongs: as a reply to the open email,
// written into the current compose window, or in a brand-new email window.
async function deliver(html: string): Promise<void> {
  if (item.value.mode === 'read') {
    status.value = displayReplyWithBody(html)
      ? t('home.boxes.emailWriting.opened')
      : t('home.boxes.emailWriting.failed')
    return
  }
  if (item.value.mode === 'compose') {
    status.value = (await setComposeBody(html))
      ? t('home.boxes.emailWriting.inserted')
      : t('home.boxes.emailWriting.failed')
    return
  }
  status.value = displayNewMessageWithBody(html)
    ? t('home.boxes.emailWriting.opened')
    : t('home.boxes.emailWriting.failed')
}
</script>

<template>
  <div class="syn-card ewb">
    <h2 class="syn-card-title">{{ t('home.boxes.emailWriting.title') }}</h2>

    <input
      ref="topicInput"
      v-model="topic"
      type="text"
      class="ewb__topic"
      :placeholder="t('home.boxes.emailWriting.topicPlaceholder')"
    />

    <div class="ewb__styles">
      <ActionButton
        v-for="style in STYLES"
        :key="style"
        class="ewb__style"
        :block="false"
        :loading="active === style"
        :disabled="active !== null && active !== style"
        @click="generate(style)"
      >
        {{ t(`tone.${style}`) }}
      </ActionButton>
    </div>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </div>
</template>

<style scoped>
.ewb {
  gap: var(--syn-space-3);
}
.ewb__topic {
  width: 100%;
}
.ewb__styles {
  display: flex;
  flex-wrap: wrap;
  gap: var(--syn-space-2);
}
.ewb__style {
  flex: 1 1 auto;
  justify-content: center;
  min-width: 88px;
}
</style>
