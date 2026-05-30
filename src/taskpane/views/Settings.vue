<script setup lang="ts">
import { onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { signOut, signedInEmail, signedInBaseUrl, isSignedIn } from '@/taskpane/composables/useAuth'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { go } from '@/taskpane/router'
import { loadSettings, patchSettings } from '@/taskpane/composables/useRoamingSettings'
import { setLocale, detectLocale, SUPPORTED_LOCALES, type Locale } from '@/i18n'
import type { ModelConfig } from '@shared/types'

const { t } = useI18n()
const { call } = useSynaplanClient()
const error = ref<string | null>(null)
const baseUrlEdit = ref<string>(signedInBaseUrl.value ?? 'https://web.synaplan.com')
const editable = ref(!isSignedIn.value)

const models = ref<ModelConfig | null>(null)
const modelsLoading = ref(false)
const modelsError = ref<string | null>(null)

function modelLabel(choice: { name: string; service?: string } | null): string {
  if (!choice) return '—'
  return choice.service ? `${choice.name} (${choice.service})` : choice.name
}

function configUrl(): string {
  const base = (signedInBaseUrl.value ?? 'https://web.synaplan.com').replace(/\/$/, '')
  return `${base}/config/ai-models`
}

function openModelConfig(): void {
  const url = configUrl()
  const ui = (typeof Office !== 'undefined' ? Office.context?.ui : undefined) as
    | { openBrowserWindow?: (u: string) => void }
    | undefined
  if (ui && typeof ui.openBrowserWindow === 'function') {
    ui.openBrowserWindow(url)
  } else {
    window.open(url, '_blank', 'noopener')
  }
}

onMounted(async () => {
  if (!isSignedIn.value) return
  modelsLoading.value = true
  modelsError.value = null
  try {
    models.value = await call((c) => c.getModelConfig())
  } catch (err) {
    modelsError.value = err instanceof Error ? err.message : String(err)
  } finally {
    modelsLoading.value = false
  }
})

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

type LanguagePref = 'auto' | Locale
const LANGUAGE_OPTIONS = ['auto', ...SUPPORTED_LOCALES] as const
const languagePref = ref<LanguagePref>((loadSettings()?.language as LanguagePref) ?? 'auto')

async function onLanguageChange(): Promise<void> {
  const pref = languagePref.value
  // Apply to the live UI immediately; 'auto' re-resolves the Outlook language.
  setLocale(pref === 'auto' ? detectLocale() : pref)
  // Persist for next launch. Roaming settings require a signed-in record;
  // when signed out the choice still applies for this session.
  if (!isSignedIn.value) return
  try {
    await patchSettings({ language: pref })
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  }
}
</script>

<template>
  <section class="settings">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('home')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ t('settings.title') }}</h2>
    </header>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('settings.signedInAs') }}</h3>
      <div class="settings__identity">
        <strong>{{ signedInEmail ?? '—' }}</strong>
        <span class="syn-card-sub">{{ signedInBaseUrl ?? '—' }}</span>
      </div>
      <ActionButton variant="primary" @click="handleSignOut">
        {{ t('settings.signOut') }}
      </ActionButton>
    </div>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('settings.synaplanInstance') }}</h3>
      <p class="syn-card-sub">{{ t('settings.instanceHint') }}</p>
      <input v-model="baseUrlEdit" :disabled="!editable" type="url" spellcheck="false" />
      <ActionButton v-if="editable" @click="saveBaseUrl">
        {{ t('common.save') }}
      </ActionButton>
    </div>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('settings.models.title') }}</h3>
      <p v-if="modelsLoading" class="syn-card-sub">{{ t('settings.models.loading') }}</p>
      <dl v-else class="settings__models">
        <div class="settings__model-row">
          <dt>{{ t('settings.models.chat') }}</dt>
          <dd>{{ modelLabel(models?.chat ?? null) }}</dd>
        </div>
        <div class="settings__model-row">
          <dt>{{ t('settings.models.imageGen') }}</dt>
          <dd>{{ modelLabel(models?.imageGen ?? null) }}</dd>
        </div>
        <div class="settings__model-row">
          <dt>{{ t('settings.models.vectorize') }}</dt>
          <dd>{{ modelLabel(models?.vectorize ?? null) }}</dd>
        </div>
      </dl>
      <Toast v-if="modelsError" kind="error" :message="modelsError" />
      <ActionButton @click="openModelConfig"> {{ t('settings.models.configure') }} → </ActionButton>
    </div>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('settings.preferences') }}</h3>
      <label class="settings__field">
        <span>{{ t('settings.defaultRagGroup') }}</span>
        <code>{{ lastRagGroupId || '—' }}</code>
      </label>
      <label class="settings__field">
        <span>{{ t('settings.language') }}</span>
        <select v-model="languagePref" class="settings__select" @change="onLanguageChange">
          <option v-for="l in LANGUAGE_OPTIONS" :key="l" :value="l">
            {{ t(`language.${l}`) }}
          </option>
        </select>
      </label>
    </div>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('settings.advanced') }}</h3>
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
  gap: var(--syn-space-3);
}
.settings__identity {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
  word-break: break-all;
}
input {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
}
.settings__field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--syn-space-2);
}
.settings__select {
  padding: var(--syn-space-1) var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  background: var(--syn-bg);
  color: var(--syn-text);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
}
.settings__models {
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.settings__model-row {
  display: flex;
  justify-content: space-between;
  gap: var(--syn-space-2);
  font-size: var(--syn-font-size-sm);
}
.settings__model-row dt {
  color: var(--syn-muted);
}
.settings__model-row dd {
  margin: 0;
  text-align: right;
  font-weight: 500;
}
</style>
