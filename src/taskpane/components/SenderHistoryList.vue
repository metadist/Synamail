<script setup lang="ts">
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import type { SenderHistoryResult } from '@shared/types'

interface Props {
  history: SenderHistoryResult
}

const props = defineProps<Props>()
defineEmits<{
  (e: 'summarise'): void
}>()

const { t, locale } = useI18n()

const dateFmt = computed(
  () =>
    new Intl.DateTimeFormat(locale.value, {
      day: '2-digit',
      month: 'short',
      year:
        new Date().getFullYear() ===
        new Date(props.history.items[0]?.date ?? Date.now()).getFullYear()
          ? undefined
          : '2-digit',
    }),
)

function fmt(iso: string): string {
  try {
    return dateFmt.value.format(new Date(iso))
  } catch {
    return iso
  }
}
</script>

<template>
  <section class="senders" data-testid="sender-history">
    <header class="senders__header">
      <h3 class="senders__title">
        {{ t('read.senderHistory.title', { email: history.email }) }}
      </h3>
      <span v-if="!history.fromOutlook" class="senders__badge" :title="t('common.mockMode')">
        {{ t('read.senderHistory.mockBadge') }}
      </span>
    </header>

    <p class="senders__meta syn-muted">
      {{ t('read.senderHistory.count', { n: history.total }) }}
    </p>

    <ul class="senders__list">
      <li
        v-for="item in history.items"
        :key="item.messageId ?? item.date + item.subject"
        class="senders__item"
        :class="{ 'is-unread': item.unread }"
      >
        <div class="senders__row">
          <span class="senders__subject">{{ item.subject || '—' }}</span>
          <span class="senders__date syn-muted">{{ fmt(item.date) }}</span>
        </div>
        <p class="senders__snippet syn-muted">{{ item.snippet }}</p>
      </li>
    </ul>

    <button
      v-if="history.items.length > 0"
      type="button"
      class="senders__summarise"
      @click="$emit('summarise')"
    >
      {{ t('read.senderHistory.summariseAll') }}
    </button>
  </section>
</template>

<style scoped>
.senders {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  padding: var(--syn-space-3);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  background: var(--syn-surface);
}
.senders__header {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
}
.senders__title {
  margin: 0;
  font-size: var(--syn-font-size-md);
  flex: 1;
}
.senders__badge {
  font-size: var(--syn-font-size-xs);
  padding: 2px var(--syn-space-1);
  border-radius: 999px;
  background: var(--syn-warn-bg);
  color: var(--syn-warn-fg);
  border: 1px solid var(--syn-warn-border);
}
.senders__meta {
  font-size: var(--syn-font-size-sm);
}
.senders__list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
  max-height: 18rem;
  overflow-y: auto;
}
.senders__item {
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  border: 1px solid transparent;
}
.senders__item.is-unread {
  border-color: var(--syn-border);
  background: var(--syn-bg);
}
.senders__item.is-unread .senders__subject {
  font-weight: 600;
}
.senders__row {
  display: flex;
  justify-content: space-between;
  gap: var(--syn-space-2);
  align-items: baseline;
}
.senders__subject {
  font-size: var(--syn-font-size-sm);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.senders__date {
  font-size: var(--syn-font-size-xs);
  flex-shrink: 0;
}
.senders__snippet {
  margin: 2px 0 0;
  font-size: var(--syn-font-size-xs);
  overflow: hidden;
  text-overflow: ellipsis;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
}
.senders__summarise {
  align-self: flex-start;
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  padding: var(--syn-space-1) 0;
  font-size: var(--syn-font-size-sm);
  text-decoration: underline;
  cursor: pointer;
}
</style>
