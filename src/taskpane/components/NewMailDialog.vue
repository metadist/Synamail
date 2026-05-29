<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { canDisplayNewMessage, displayNewMessage } from '@/taskpane/composables/useOutlookCompose'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

const emit = defineEmits<{ (e: 'cancel'): void; (e: 'opened'): void }>()

const { t } = useI18n()
const { call } = useSynaplanClient()

const description = ref('')
const language = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('en')
const draft = ref<{ subject: string; htmlBody: string } | null>(null)
const generating = ref(false)
const opening = ref(false)
const error = ref<string | null>(null)

const canOpen = canDisplayNewMessage()

async function generate(): Promise<void> {
  const desc = description.value.trim()
  if (!desc) return
  generating.value = true
  error.value = null
  try {
    const r = await call((c) =>
      c.composeNew({
        description: desc,
        language: language.value === 'auto' ? 'en' : language.value,
      }),
    )
    if (r) draft.value = r
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    generating.value = false
  }
}

async function open(): Promise<void> {
  if (!draft.value) return
  opening.value = true
  error.value = null
  try {
    await displayNewMessage({ subject: draft.value.subject, htmlBody: draft.value.htmlBody })
    emit('opened')
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    opening.value = false
  }
}
</script>

<template>
  <div class="nm" role="dialog" aria-modal="true">
    <div class="nm__card">
      <h3 class="nm__title">{{ t('home.newMail.title') }}</h3>

      <div class="nm__field">
        <label for="nm-desc">{{ t('home.newMail.description') }}</label>
        <textarea
          id="nm-desc"
          v-model="description"
          rows="3"
          :placeholder="t('home.newMail.descriptionPlaceholder')"
        />
      </div>

      <div class="syn-row">
        <ActionButton
          :block="false"
          :loading="generating"
          :disabled="!description.trim()"
          @click="generate"
        >
          {{ t('home.newMail.generate') }}
        </ActionButton>
        <LanguagePicker v-model="language" />
      </div>

      <div v-if="draft" class="nm__preview">
        <p class="nm__subject">
          <strong>{{ draft.subject }}</strong>
        </p>
        <!-- eslint-disable-next-line vue/no-v-html -->
        <div class="nm__body" v-html="draft.htmlBody" />
      </div>

      <Toast v-if="!canOpen" kind="info" :message="t('home.newMail.unsupportedHost')" />
      <Toast v-if="error" kind="error" :message="error" />

      <div class="nm__actions">
        <ActionButton :block="false" @click="emit('cancel')">
          {{ t('common.cancel') }}
        </ActionButton>
        <ActionButton
          variant="primary"
          :block="false"
          :loading="opening"
          :disabled="!draft || !canOpen"
          @click="open"
        >
          {{ t('home.newMail.openInOutlook') }}
        </ActionButton>
      </div>
    </div>
  </div>
</template>

<style scoped>
.nm {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.35);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: var(--syn-space-3);
  z-index: 50;
}
.nm__card {
  background: var(--syn-bg);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  padding: var(--syn-space-4);
  width: 100%;
  max-width: 380px;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
  max-height: 90vh;
  overflow-y: auto;
}
.nm__title {
  margin: 0;
  font-size: var(--syn-font-size-lg);
}
.nm__field {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.nm__field label {
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
}
.nm__field textarea {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
  resize: vertical;
}
.nm__preview {
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  padding: var(--syn-space-2);
  background: var(--syn-surface);
  font-size: var(--syn-font-size-sm);
}
.nm__subject {
  margin: 0 0 var(--syn-space-2);
}
.nm__body {
  white-space: normal;
}
.nm__actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--syn-space-2);
}
</style>
