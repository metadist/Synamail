<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { signOut, signedInEmail, signedInBaseUrl, isSignedIn } from '@/taskpane/composables/useAuth'
import { go } from '@/taskpane/router'
import { loadSettings, patchSettings } from '@/taskpane/composables/useRoamingSettings'

const { t } = useI18n()
const error = ref<string | null>(null)
const baseUrlEdit = ref<string>(signedInBaseUrl.value ?? 'https://web.synaplan.com')
const editable = ref(!isSignedIn.value)

async function handleSignOut(): Promise<void> {
  error.value = null
  try {
    await signOut({ revokeRemote: true })
    go('sign-in')
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  }
}

async function saveBaseUrl(): Promise<void> {
  if (!isSignedIn.value) {
    // Pre-signin override is consumed by SignIn.vue's ref.
    return
  }
  try {
    await patchSettings({ baseUrl: baseUrlEdit.value })
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  }
}

const lastRagGroupId = loadSettings()?.lastRagGroupId ?? ''
const language = loadSettings()?.language ?? 'auto'
</script>

<template>
  <section class="settings">
    <header class="syn-row">
      <button type="button" class="settings__back" @click="go('read')">
        ← {{ t('common.back') }}
      </button>
      <h2>{{ t('settings.title') }}</h2>
    </header>

    <div class="settings__section">
      <p class="syn-muted">
        {{ t('settings.signedInAs') }}
      </p>
      <p>
        <strong>{{ signedInEmail ?? '—' }}</strong>
      </p>
      <p class="syn-muted">on {{ signedInBaseUrl ?? '—' }}</p>
      <ActionButton variant="primary" @click="handleSignOut">
        {{ t('settings.signOut') }}
      </ActionButton>
    </div>

    <div class="settings__section">
      <p class="syn-muted">
        {{ t('settings.synaplanInstance') }}
      </p>
      <p class="syn-muted">
        {{ t('settings.instanceHint') }}
      </p>
      <input v-model="baseUrlEdit" :disabled="!editable" type="url" spellcheck="false" />
      <ActionButton v-if="editable" @click="saveBaseUrl">
        {{ t('common.save') }}
      </ActionButton>
    </div>

    <div class="settings__section">
      <p>
        <strong>{{ t('settings.preferences') }}</strong>
      </p>
      <label
        >{{ t('settings.defaultRagGroup') }}: <code>{{ lastRagGroupId || '—' }}</code></label
      >
      <label
        >{{ t('settings.language') }}: <code>{{ language }}</code></label
      >
    </div>

    <div class="settings__section">
      <p>
        <strong>{{ t('settings.advanced') }}</strong>
      </p>
      <ActionButton @click="go('rule-editor')"> {{ t('settings.routingRules') }} → </ActionButton>
    </div>

    <Toast v-if="error" kind="error" :message="error" />
  </section>
</template>

<style scoped>
.settings {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-4);
}
.settings__back {
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  padding: 0;
}
.settings__section {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  padding: var(--syn-space-3);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
}
input {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
}
</style>
