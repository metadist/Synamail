<script setup lang="ts">
/**
 * Compact profiling surface for the Home "Profiling" accordion.
 *
 * Goal: when the flap opens, the most useful thing (a short profile extract)
 * and the primary action ("Update the profile") are both above the fold — no
 * scrolling, no detour through a separate view. The full profile (facts, open
 * loops, search, ask) still lives in ContactProfile.vue, one tap away.
 */
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useContactCounterpart } from '@/taskpane/composables/useContactCounterpart'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage, isApiError } from '@shared/synaplan-client'
import { openContactProfile } from '@/taskpane/router'
import type { ContactProfileData } from '@shared/types'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()
const { contactEmail, direction } = useContactCounterpart(item)

const profile = ref<ContactProfileData | null>(null)
const profileLoaded = ref(false)
const profilingUnavailable = ref(false)
const active = ref<string | null>(null)
const error = ref<string | null>(null)
const status = ref<string | null>(null)

const emailOpen = computed(() => item.value.mode === 'read')
const canRoll = computed(() => emailOpen.value && !!item.value.bodyText && !!contactEmail.value)

async function run<T>(key: string, fn: () => Promise<T | null>): Promise<T | null> {
  active.value = key
  error.value = null
  try {
    return await fn()
  } catch (err) {
    if (isApiError(err) && err.code === 'PROFILING_UNAVAILABLE') {
      profilingUnavailable.value = true
      return null
    }
    error.value = errorMessage(err)
    return null
  } finally {
    active.value = null
  }
}

async function load(): Promise<void> {
  status.value = null
  profile.value = null
  profileLoaded.value = false
  if (!contactEmail.value) return
  await run('load', async () => {
    profile.value = await call((c) => c.getContactProfile(contactEmail.value))
    profileLoaded.value = true
    return profile.value
  })
}

async function update(): Promise<void> {
  if (!canRoll.value) return
  status.value = null
  const r = await run('update', () =>
    call((c) =>
      c.updateContactProfile({
        email: contactEmail.value,
        subject: item.value.subject,
        body: item.value.bodyText,
        direction: direction.value,
      }),
    ),
  )
  if (r) {
    profile.value = r
    profileLoaded.value = true
    status.value = t('contactProfile.profile.updated')
  }
}

function asOfLine(p: ContactProfileData): string {
  const date = p.updatedAt ? new Date(p.updatedAt).toLocaleDateString() : '—'
  return t('contactProfile.profile.asOf', { date, n: p.emailCount })
}

// Reload whenever the open email (and thus the contact) changes.
watch(contactEmail, load, { immediate: true })
</script>

<template>
  <div class="pp">
    <p v-if="!emailOpen" class="syn-muted">{{ t('read.noEmail') }}</p>
    <p v-else-if="!contactEmail" class="syn-muted">{{ t('contactProfile.noContact') }}</p>

    <template v-else>
      <p v-if="profilingUnavailable" class="syn-muted">
        {{ t('contactProfile.profile.unavailable') }}
      </p>

      <template v-else>
        <!-- Primary action first, so it's reachable the instant the flap opens. -->
        <div class="pp__actions">
          <ActionButton
            variant="primary"
            :block="false"
            :disabled="!canRoll"
            :loading="active === 'update'"
            @click="update"
          >
            {{ t('contactProfile.profile.update') }}
          </ActionButton>
          <button type="button" class="pp__link" @click="openContactProfile(contactEmail)">
            {{ t('contactProfile.openFull') }} →
          </button>
        </div>

        <p class="pp__contact syn-muted">
          {{ contactEmail }}<template v-if="profile"> · {{ asOfLine(profile) }}</template>
        </p>

        <p v-if="!canRoll" class="pp__hint syn-muted">
          {{ t('contactProfile.profile.noEmailOpen') }}
        </p>

        <Toast v-if="status" kind="success" :message="status" />
        <Toast v-if="error" kind="error" :message="error" />

        <p v-if="!profileLoaded" class="syn-muted">{{ t('common.loading') }}</p>
        <!-- Short extract only — the full narrative is on the dedicated view. -->
        <p v-else-if="profile" class="pp__summary">{{ profile.summary }}</p>
        <p v-else class="syn-muted">{{ t('contactProfile.profile.empty') }}</p>
      </template>
    </template>
  </div>
</template>

<style scoped>
.pp {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.pp__actions {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
  flex-wrap: wrap;
}
/* Quiet text link to the full profile view — secondary to "Update". */
.pp__link {
  background: none;
  border: 0;
  padding: 0;
  font: inherit;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-brand-600);
  cursor: pointer;
  text-decoration: underline;
  margin-left: auto;
}
.pp__contact {
  margin: 0;
  font-size: var(--syn-font-size-sm);
  overflow-wrap: anywhere;
}
.pp__hint {
  margin: 0;
  font-size: var(--syn-font-size-sm);
}
/* Keep the extract short so the flap stays above the fold; the full summary is
   one tap away on the dedicated profile view. */
.pp__summary {
  margin: 0;
  font-size: var(--syn-font-size-sm);
  white-space: pre-wrap;
  display: -webkit-box;
  -webkit-line-clamp: 4;
  line-clamp: 4;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
</style>
