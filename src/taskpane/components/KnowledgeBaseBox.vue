<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import SaveToRagDialog from '@/taskpane/components/SaveToRagDialog.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { getReadItemAsFile, useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { getLastRagGroupId, setLastRagGroupId } from '@/taskpane/composables/useRoamingSettings'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import type { RagGroup } from '@shared/types'

type ProcessLevel = 'store' | 'extract' | 'vectorize' | 'full'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

const groups = ref<RagGroup[]>([])
const groupsLoaded = ref(false)
// The configured "standard" folder (last folder the user saved to).
const standardId = ref<string>(getLastRagGroupId() ?? '')
const active = ref<string | null>(null)
const showDialog = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

const isRead = computed(() => item.value.mode === 'read')
const senderEmail = computed(() => item.value.from ?? '')

// Up to three existing folders offered as one-tap topics, excluding the
// standard folder (which already has its own primary button).
const topicButtons = computed(() =>
  groups.value.filter((g) => g.id && g.id !== standardId.value).slice(0, 3),
)

function groupLabel(id: string): string {
  return groups.value.find((g) => g.id === id)?.name || id
}

// Fetch folders once, the first time an email is open (no email → nothing to
// save, so we avoid the round-trip entirely).
watch(
  isRead,
  async (read) => {
    if (!read || groupsLoaded.value) return
    groupsLoaded.value = true
    try {
      groups.value = (await call((c) => c.ragGroups())) ?? []
    } catch (err) {
      error.value = errorMessage(err)
    }
  },
  { immediate: true },
)

async function save(groupId: string, processLevel: ProcessLevel = 'extract'): Promise<void> {
  active.value = groupId
  error.value = null
  status.value = null
  try {
    const file = getReadItemAsFile(item.value)
    const upload = (level: ProcessLevel) =>
      call((c) =>
        c.fileUpload({
          filename: file.filename,
          contentBase64: file.contentBase64,
          mimeType: file.mimeType,
          groupId,
          metadata: {
            from: item.value.from ?? '',
            subject: item.value.subject,
            to: item.value.to.join(', '),
          },
          processLevel: level,
        }),
      )

    // `vectorize`/`full` can fail server-side without an embedding model; retry
    // once at `extract` so the email is still saved rather than lost.
    let r: Awaited<ReturnType<typeof upload>>
    if (processLevel === 'vectorize' || processLevel === 'full') {
      try {
        r = await upload(processLevel)
      } catch {
        r = await upload('extract')
      }
    } else {
      r = await upload(processLevel)
    }

    if (r) {
      standardId.value = groupId
      if (!groups.value.some((g) => g.id === groupId)) {
        groups.value = [{ id: groupId, name: groupId }, ...groups.value]
      }
      try {
        await setLastRagGroupId(groupId)
      } catch {
        // Roaming write can fail offline/in tests; the in-memory state still holds.
      }
      status.value = t('home.boxes.knowledge.saved', { group: groupLabel(groupId) })
    }
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    active.value = null
  }
}

function openDialog(): void {
  error.value = null
  status.value = null
  showDialog.value = true
}

async function handleSaveConfirm(payload: {
  groupId: string
  processLevel: ProcessLevel
}): Promise<void> {
  showDialog.value = false
  await save(payload.groupId, payload.processLevel)
}
</script>

<template>
  <div class="syn-card kb">
    <h2 class="syn-card-title">{{ t('home.boxes.knowledge.title') }}</h2>

    <p v-if="!isRead" class="syn-muted">{{ t('home.boxes.noEmail') }}</p>

    <template v-else>
      <div class="kb__buttons">
        <!-- Standard folder if the user has one, otherwise a "New" entry that
             opens the full save dialog. -->
        <ActionButton
          v-if="standardId"
          class="kb__btn"
          variant="primary"
          :block="false"
          :loading="active === standardId"
          :disabled="active !== null && active !== standardId"
          @click="save(standardId)"
        >
          {{ groupLabel(standardId) }}
        </ActionButton>
        <ActionButton v-else class="kb__btn" variant="primary" :block="false" @click="openDialog">
          {{ t('home.boxes.knowledge.new') }}
        </ActionButton>

        <!-- Up to three existing folders as quick topics. -->
        <ActionButton
          v-for="g in topicButtons"
          :key="g.id"
          class="kb__btn"
          :block="false"
          :loading="active === g.id"
          :disabled="active !== null && active !== g.id"
          @click="save(g.id)"
        >
          {{ g.name }}
        </ActionButton>

        <!-- When a standard folder exists, keep "New" available too. -->
        <ActionButton v-if="standardId" class="kb__btn" :block="false" @click="openDialog">
          {{ t('home.boxes.knowledge.new') }}
        </ActionButton>
      </div>

      <Toast v-if="status" kind="success" :message="status" />
      <Toast v-if="error" kind="error" :message="error" />
    </template>

    <SaveToRagDialog
      v-if="showDialog"
      :contact-email="senderEmail"
      :last-used-group-id="standardId"
      @cancel="showDialog = false"
      @confirm="handleSaveConfirm"
    />
  </div>
</template>

<style scoped>
.kb {
  gap: var(--syn-space-3);
}
.kb__buttons {
  display: flex;
  flex-wrap: wrap;
  gap: var(--syn-space-2);
}
.kb__btn {
  flex: 1 1 auto;
  justify-content: center;
  min-width: 88px;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
</style>
