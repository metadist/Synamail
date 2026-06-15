<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { signIn, defaultBaseUrl } from '@/taskpane/composables/useAuth'
import { errorMessage } from '@shared/synaplan-client'
import { setPreferredBaseUrl } from '@/taskpane/composables/useRoamingSettings'
import { go } from '@/taskpane/router'

const { t } = useI18n()
const loading = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const savedUrl = ref<string>(defaultBaseUrl())
const baseUrlOverride = ref<string>(savedUrl.value)

function normalize(url: string): string {
  return url.trim().replace(/\/+$/, '')
}

/** Returns an error message key, or null when the URL is usable. */
function validate(url: string): string | null {
  const v = normalize(url)
  let parsed: URL
  try {
    parsed = new URL(v)
  } catch {
    return t('signIn.invalidUrl')
  }
  // Office's sign-in dialog (displayDialogAsync) hard-rejects non-HTTPS URLs,
  // so an http://localhost instance can't be used directly — point the user
  // at the local HTTPS bridge instead.
  if (parsed.protocol !== 'https:') return t('signIn.httpsRequired')
  return null
}

const dirty = computed(() => normalize(baseUrlOverride.value) !== savedUrl.value)

/** Registration page on the chosen instance (falls back to the default). */
const registerUrl = computed(() => {
  const base = validate(baseUrlOverride.value) ? defaultBaseUrl() : normalize(baseUrlOverride.value)
  return `${base}/register`
})

async function saveInstance(): Promise<void> {
  error.value = null
  status.value = null
  const invalid = validate(baseUrlOverride.value)
  if (invalid) {
    error.value = invalid
    return
  }
  const v = normalize(baseUrlOverride.value)
  try {
    await setPreferredBaseUrl(v)
    baseUrlOverride.value = v
    savedUrl.value = v
    status.value = t('signIn.instanceSaved', { url: v })
  } catch (err) {
    error.value = errorMessage(err)
  }
}

async function handleClick(): Promise<void> {
  error.value = null
  status.value = null
  const invalid = validate(baseUrlOverride.value)
  if (invalid) {
    error.value = invalid
    return
  }
  loading.value = true
  try {
    const v = normalize(baseUrlOverride.value)
    // Remember the instance even if the round-trip fails, so the field keeps
    // the user's choice next time.
    await setPreferredBaseUrl(v).catch(() => undefined)
    savedUrl.value = v
    await signIn({ baseUrl: v })
    go('home')
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <section class="signin">
    <img src="/assets/icon-64.png" alt="" width="48" height="48" class="signin__logo" />
    <h1 class="signin__title">
      {{ t('signIn.title') }}
    </h1>
    <p class="syn-muted signin__subtitle">
      {{ t('signIn.subtitle') }}
    </p>

    <div class="signin__override">
      <label for="baseUrl" class="syn-muted">{{ t('settings.synaplanInstance') }}</label>
      <input
        id="baseUrl"
        v-model="baseUrlOverride"
        type="url"
        spellcheck="false"
        autocomplete="off"
        @keyup.enter="saveInstance"
      />
      <div class="signin__override-actions">
        <ActionButton v-if="dirty" :block="false" @click="saveInstance">
          {{ t('common.save') }}
        </ActionButton>
      </div>
      <p class="syn-muted signin__hint">{{ t('signIn.localDevHint') }}</p>
    </div>

    <ActionButton variant="primary" :loading="loading" @click="handleClick">
      {{ loading ? t('signIn.loading') : t('signIn.button') }}
    </ActionButton>

    <p class="syn-muted signin__create">
      {{ t('signIn.noAccount') }}
      <a :href="registerUrl" target="_blank" rel="noopener noreferrer">{{
        t('signIn.createAccount')
      }}</a>
    </p>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </section>
</template>

<style scoped>
.signin {
  padding: var(--syn-space-6) var(--syn-space-4);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
  align-items: stretch;
}
.signin__logo {
  align-self: center;
}
.signin__title {
  margin: 0;
  font-size: var(--syn-font-size-xl);
  text-align: center;
}
.signin__subtitle {
  margin: 0 0 var(--syn-space-3);
  text-align: center;
}
.signin__override {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.signin__override-actions {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
}
.signin__hint {
  margin: 0;
  font-size: var(--syn-font-size-sm);
}
.signin__create {
  margin: var(--syn-space-1) 0 0;
  text-align: center;
  font-size: var(--syn-font-size-sm);
}
</style>
