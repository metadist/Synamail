<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import TonePicker from '@/taskpane/components/TonePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { setComposeBody, useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

const composeOpen = computed(() => item.value.mode === 'compose')
const intent = ref('')
const tone = ref<'formal' | 'concise' | 'friendly'>('concise')
const targetLang = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('auto')
const loading = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

async function draft(): Promise<void> {
  const text = intent.value.trim()
  if (!text) return
  loading.value = true
  error.value = null
  status.value = null
  try {
    const r = await call((c) =>
      c.composeDraft({
        intent: text,
        tone: tone.value,
        language: targetLang.value === 'auto' ? 'en' : targetLang.value,
        // On a reply/forward the draft already carries the quoted original —
        // pass it so the AI can respond in context.
        referenceBody: item.value.bodyText || undefined,
      }),
    )
    if (r) {
      const ok = await setComposeBody(r.htmlBody)
      if (ok) status.value = t('compose.inserted')
      else error.value = t('compose.insertFailed')
    }
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <div class="ca">
    <p v-if="!composeOpen" class="syn-muted">{{ t('compose.noCompose') }}</p>

    <template v-else>
      <Toast v-if="status" kind="success" :message="status" />
      <Toast v-if="error" kind="error" :message="error" />

      <p class="syn-card-sub">{{ t('compose.hint') }}</p>

      <textarea
        v-model="intent"
        class="ca__intent"
        rows="3"
        :placeholder="t('compose.placeholder')"
        @keyup.enter.ctrl="draft"
      />

      <div class="ca__row">
        <TonePicker v-model="tone" />
        <LanguagePicker v-model="targetLang" />
        <ActionButton
          class="ca__go"
          variant="primary"
          :block="false"
          :loading="loading"
          :disabled="!intent.trim()"
          @click="draft"
        >
          {{ t('compose.draft') }}
        </ActionButton>
      </div>
    </template>
  </div>
</template>

<style scoped>
.ca {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.ca__intent {
  width: 100%;
  resize: vertical;
}
.ca__row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: var(--syn-space-2);
}
.ca__go {
  min-width: 72px;
  justify-content: center;
}
</style>
