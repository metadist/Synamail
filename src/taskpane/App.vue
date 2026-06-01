<script setup lang="ts">
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import SignIn from './views/SignIn.vue'
import Home from './views/Home.vue'
import ReadMode from './views/ReadMode.vue'
import ComposeMode from './views/ComposeMode.vue'
import Settings from './views/Settings.vue'
import RuleEditor from './views/RuleEditor.vue'
import MailRoutes from './views/MailRoutes.vue'
import ContactKnowledgeBase from './views/ContactKnowledgeBase.vue'
import { currentView, go } from './router'
import { isSignedIn } from './composables/useAuth'

const { t } = useI18n()
const view = computed(() => currentView.value)

const components = {
  'sign-in': SignIn,
  home: Home,
  read: ReadMode,
  compose: ComposeMode,
  settings: Settings,
  'rule-editor': RuleEditor,
  'mail-routes': MailRoutes,
  'contact-kb': ContactKnowledgeBase,
} as const
</script>

<template>
  <div class="app">
    <header class="app__bar">
      <button
        v-if="isSignedIn"
        type="button"
        class="app__brand"
        :aria-label="t('home.title')"
        @click="go('home')"
      >
        <strong>Synamail</strong>
      </button>
      <strong v-else>Synamail</strong>
      <nav v-if="isSignedIn" class="app__nav">
        <button
          type="button"
          class="app__icon"
          :class="{ 'app__icon--active': view === 'home' }"
          :aria-label="t('home.title')"
          @click="go('home')"
        >
          ⌂
        </button>
        <button
          type="button"
          class="app__icon"
          :class="{ 'app__icon--active': view === 'settings' }"
          :aria-label="t('settings.title')"
          @click="go('settings')"
        >
          ⚙
        </button>
      </nav>
    </header>
    <component :is="components[view]" />
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}
.app__bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: var(--syn-space-2) var(--syn-space-3);
  border-bottom: 1px solid var(--syn-border);
  background: var(--syn-surface);
}
.app__brand {
  background: none;
  border: 0;
  color: var(--syn-text);
  padding: 0;
  cursor: pointer;
  font: inherit;
}
.app__nav {
  display: flex;
  gap: var(--syn-space-1);
}
.app__icon {
  background: none;
  border: 0;
  font-size: 1.1rem;
  color: var(--syn-muted);
  padding: var(--syn-space-1) var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  cursor: pointer;
  line-height: 1;
}
.app__icon--active {
  color: var(--syn-text);
  background: var(--syn-bg);
}
</style>
