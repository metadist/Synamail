<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { signIn, defaultBaseUrl } from '@/taskpane/composables/useAuth'
import { go } from '@/taskpane/router'

const { t } = useI18n()
const loading = ref(false)
const error = ref<string | null>(null)
const baseUrlOverride = ref<string>(defaultBaseUrl())
const showOverride = ref(false)

async function handleClick(): Promise<void> {
  loading.value = true
  error.value = null
  try {
    await signIn({ baseUrl: baseUrlOverride.value })
    go('read')
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
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

    <ActionButton variant="primary" :loading="loading" @click="handleClick">
      {{ loading ? t('signIn.loading') : t('signIn.button') }}
    </ActionButton>

    <button class="signin__link" type="button" @click="showOverride = !showOverride">
      {{ t('signIn.selfHosted') }} →
    </button>

    <div v-if="showOverride" class="signin__override">
      <label for="baseUrl" class="syn-muted">{{ t('settings.synaplanInstance') }}</label>
      <input
        id="baseUrl"
        v-model="baseUrlOverride"
        type="url"
        spellcheck="false"
        autocomplete="off"
      />
    </div>

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
.signin__link {
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  text-decoration: underline;
  padding: var(--syn-space-1);
  align-self: center;
}
.signin__override {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.signin__override input {
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  border: 1px solid var(--syn-border);
  font-family: inherit;
}
</style>
