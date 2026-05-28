<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'

interface Props {
  senderEmail: string
}

defineProps<Props>()

const emit = defineEmits<{
  (e: 'cancel'): void
  (e: 'confirm', payload: { alsoCleanExisting: boolean }): void
}>()

const { t } = useI18n()
const alsoCleanExisting = ref(true)
</script>

<template>
  <div class="block-dialog" role="dialog" aria-modal="true">
    <div class="block-dialog__card">
      <h3 class="block-dialog__title">{{ t('read.blockDialog.title') }}</h3>

      <p class="syn-muted">
        {{ t('read.blockDialog.body', { email: senderEmail }) }}
      </p>

      <label class="block-dialog__check">
        <input v-model="alsoCleanExisting" type="checkbox" />
        <span>{{ t('read.blockDialog.alsoClean') }}</span>
      </label>

      <p class="block-dialog__note syn-muted">
        {{ t('read.blockDialog.undoHint') }}
      </p>

      <div class="block-dialog__actions">
        <ActionButton :block="false" @click="emit('cancel')">
          {{ t('common.cancel') }}
        </ActionButton>
        <ActionButton
          variant="primary"
          :block="false"
          @click="emit('confirm', { alsoCleanExisting })"
        >
          {{ t('read.blockDialog.confirm') }}
        </ActionButton>
      </div>
    </div>
  </div>
</template>

<style scoped>
.block-dialog {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.35);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: var(--syn-space-3);
  z-index: 50;
}
.block-dialog__card {
  background: var(--syn-bg);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  padding: var(--syn-space-4);
  width: 100%;
  max-width: 360px;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.block-dialog__title {
  margin: 0;
  font-size: var(--syn-font-size-lg);
}
.block-dialog__check {
  display: flex;
  gap: var(--syn-space-2);
  align-items: flex-start;
  font-size: var(--syn-font-size-sm);
}
.block-dialog__note {
  font-size: var(--syn-font-size-sm);
}
.block-dialog__actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--syn-space-2);
}
</style>
