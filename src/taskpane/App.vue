<script setup lang="ts">
import { computed } from 'vue'
import SignIn from './views/SignIn.vue'
import ReadMode from './views/ReadMode.vue'
import ComposeMode from './views/ComposeMode.vue'
import Settings from './views/Settings.vue'
import RuleEditor from './views/RuleEditor.vue'
import ContactKnowledgeBase from './views/ContactKnowledgeBase.vue'
import { currentView, go } from './router'
import { isSignedIn } from './composables/useAuth'

const view = computed(() => currentView.value)

const components = {
  'sign-in': SignIn,
  read: ReadMode,
  compose: ComposeMode,
  settings: Settings,
  'rule-editor': RuleEditor,
  'contact-kb': ContactKnowledgeBase,
} as const
</script>

<template>
  <div class="app">
    <header class="app__bar">
      <strong>Synamail</strong>
      <button
        v-if="isSignedIn"
        type="button"
        class="app__settings"
        :aria-label="'Settings'"
        @click="go('settings')"
      >
        ⚙
      </button>
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
.app__settings {
  background: none;
  border: 0;
  font-size: 1.1rem;
  color: var(--syn-text);
  padding: var(--syn-space-1);
  cursor: pointer;
}
</style>
