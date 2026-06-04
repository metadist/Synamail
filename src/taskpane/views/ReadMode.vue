<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import EmailActionsPanel from '@/taskpane/components/EmailActionsPanel.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { go } from '@/taskpane/router'

const { t } = useI18n()
const { item } = useOutlookItem()
</script>

<template>
  <section class="read">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('home')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ item.subject || '—' }}</h2>
      <p v-if="item.from" class="syn-view-subtitle">{{ t('read.from') }}: {{ item.from }}</p>
    </header>

    <EmailActionsPanel />
  </section>
</template>

<style scoped>
.read {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
</style>
