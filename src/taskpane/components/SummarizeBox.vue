<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { displayNewMessageWithBody, useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { summaryLanguageOptions } from '@/taskpane/composables/useLanguagePrefs'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { renderMarkdown } from '@shared/markdown'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

// Up to three language buttons: configured standard language first, then
// German and English when they differ from the standard.
const languages = summaryLanguageOptions()

const active = ref<string | null>(null)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

const hasEmail = computed(() => item.value.bodyText.trim().length > 0)

async function summarise(lang: string): Promise<void> {
  active.value = lang
  error.value = null
  status.value = null
  try {
    const r = await call((c) =>
      c.summarise({
        subject: item.value.subject,
        body: item.value.bodyText,
        from: item.value.from,
        to: item.value.to,
        language: lang,
      }),
    )
    if (r) {
      const subject = item.value.subject
        ? `${t('read.actions.summarise')}: ${item.value.subject}`
        : t('read.actions.summarise')
      status.value = displayNewMessageWithBody(renderMarkdown(r.summary), subject)
        ? t('home.boxes.summarize.opened')
        : t('home.boxes.summarize.failed')
    }
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    active.value = null
  }
}
</script>

<template>
  <div class="syn-card sb">
    <h2 class="syn-card-title">{{ t('home.boxes.summarize.title') }}</h2>

    <p v-if="!hasEmail" class="syn-muted">{{ t('home.boxes.noEmail') }}</p>

    <div v-else class="sb__langs">
      <ActionButton
        v-for="lang in languages"
        :key="lang"
        class="sb__lang"
        :block="false"
        :loading="active === lang"
        :disabled="active !== null && active !== lang"
        @click="summarise(lang)"
      >
        {{ t(`language.${lang}`) }}
      </ActionButton>
    </div>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </div>
</template>

<style scoped>
.sb {
  gap: var(--syn-space-3);
}
.sb__langs {
  display: flex;
  flex-wrap: wrap;
  gap: var(--syn-space-2);
}
.sb__lang {
  flex: 1 1 auto;
  justify-content: center;
  min-width: 88px;
}
</style>
