<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import type { RagGroup } from '@shared/types'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'

interface Props {
  /** Lower-cased sender email used to suggest a `contact:<email>` group. */
  contactEmail?: string
  /** Group id chosen the last time the user saved (pre-selected on mount). */
  lastUsedGroupId?: string
}

const props = defineProps<Props>()

const emit = defineEmits<{
  (e: 'cancel'): void
  (e: 'confirm', payload: { groupId: string; processLevel: ProcessLevel }): void
}>()

type ProcessLevel = 'store' | 'extract' | 'vectorize' | 'full'

const { t } = useI18n()
const { call } = useSynaplanClient()

const groups = ref<RagGroup[]>([])
const loadingGroups = ref(false)
const loadError = ref<string | null>(null)
const groupId = ref<string>('')
const newGroupName = ref('')
// Default to `extract` (store + text extraction): it has no embedding-model
// dependency, so a save always succeeds. `vectorize`/`full` remain explicit
// opt-ins for full RAG indexing.
const processLevel = ref<ProcessLevel>('extract')

const contactGroupId = computed(() =>
  props.contactEmail ? `contact:${props.contactEmail.toLowerCase()}` : '',
)

const combinedGroups = computed<RagGroup[]>(() => {
  const list = [...groups.value]
  const suggestions: RagGroup[] = []
  if (contactGroupId.value && !list.some((g) => g.id === contactGroupId.value)) {
    suggestions.push({
      id: contactGroupId.value,
      name: contactGroupId.value,
      description: t('read.saveDialog.contactGroupHint'),
    })
  }
  return [...suggestions, ...list]
})

onMounted(async () => {
  loadingGroups.value = true
  loadError.value = null
  try {
    const list = await call((c) => c.ragGroups())
    groups.value = list ?? []
    const preferred =
      props.lastUsedGroupId || contactGroupId.value || combinedGroups.value[0]?.id || ''
    if (preferred) groupId.value = preferred
  } catch (err) {
    loadError.value = errorMessage(err)
  } finally {
    loadingGroups.value = false
  }
})

function confirm(): void {
  const id = newGroupName.value.trim() || groupId.value.trim()
  if (!id) return
  emit('confirm', { groupId: id, processLevel: processLevel.value })
}
</script>

<template>
  <div class="save-dialog" role="dialog" aria-modal="true">
    <div class="save-dialog__card">
      <h3 class="save-dialog__title">{{ t('read.saveDialog.title') }}</h3>

      <div class="save-dialog__field">
        <label for="rag-group">{{ t('read.saveDialog.group') }}</label>
        <select id="rag-group" v-model="groupId" :disabled="loadingGroups">
          <option v-if="loadingGroups" value="">{{ t('read.saveDialog.loadingGroups') }}</option>
          <option v-for="g in combinedGroups" :key="g.id" :value="g.id">
            {{ g.name }}<span v-if="g.description"> — {{ g.description }}</span>
          </option>
        </select>
      </div>

      <div class="save-dialog__field">
        <label for="rag-new">{{ t('read.saveDialog.newGroup') }}</label>
        <input
          id="rag-new"
          v-model="newGroupName"
          type="text"
          :placeholder="t('read.saveDialog.newGroupPlaceholder')"
        />
      </div>

      <div class="save-dialog__field">
        <label for="rag-level">{{ t('read.saveDialog.processLevel') }}</label>
        <select id="rag-level" v-model="processLevel">
          <option value="store">{{ t('read.saveDialog.processLevels.store') }}</option>
          <option value="extract">{{ t('read.saveDialog.processLevels.extract') }}</option>
          <option value="vectorize">
            {{ t('read.saveDialog.processLevels.vectorize') }}
          </option>
          <option value="full">{{ t('read.saveDialog.processLevels.full') }}</option>
        </select>
      </div>

      <Toast v-if="loadError" kind="error" :message="loadError" />

      <div class="save-dialog__actions">
        <ActionButton :block="false" @click="emit('cancel')">
          {{ t('common.cancel') }}
        </ActionButton>
        <ActionButton
          variant="primary"
          :block="false"
          :disabled="!newGroupName.trim() && !groupId.trim()"
          @click="confirm"
        >
          {{ t('common.save') }}
        </ActionButton>
      </div>
    </div>
  </div>
</template>

<style scoped>
.save-dialog {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.35);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: var(--syn-space-3);
  z-index: 50;
}
.save-dialog__card {
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
.save-dialog__title {
  margin: 0;
  font-size: var(--syn-font-size-lg);
}
.save-dialog__field {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.save-dialog__field label {
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
}
.save-dialog__field select,
.save-dialog__field input {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  background: var(--syn-surface);
  color: var(--syn-text);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
}
.save-dialog__actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--syn-space-2);
}
</style>
