<script setup lang="ts">
import { onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { go } from '@/taskpane/router'
import type { RoutingCandidate, RoutingTopic } from '@shared/types'

const { t } = useI18n()
const { call } = useSynaplanClient()
const { item } = useOutlookItem()

const topics = ref<RoutingTopic[]>([])
const loading = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const savingId = ref<number | null>(null)

// Local edit buffers keyed by topic id, so each row edits independently.
const editRules = ref<Record<number, string>>({})
const editKeywords = ref<Record<number, string>>({})

const candidates = ref<RoutingCandidate[] | null>(null)
const testing = ref(false)

async function load(): Promise<void> {
  loading.value = true
  error.value = null
  try {
    const r = await call((c) => c.listTopics())
    if (r) {
      topics.value = r
      for (const tp of r) {
        editRules.value[tp.id] = tp.selectionRules ?? ''
        editKeywords.value[tp.id] = tp.keywords ?? ''
      }
    }
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  } finally {
    loading.value = false
  }
}

async function saveTopic(tp: RoutingTopic): Promise<void> {
  error.value = null
  status.value = null
  savingId.value = tp.id
  try {
    const updated = await call((c) =>
      c.updateTopicRules(tp.id, {
        selectionRules: editRules.value[tp.id] || null,
        keywords: editKeywords.value[tp.id] || null,
      }),
    )
    if (updated) {
      const idx = topics.value.findIndex((x) => x.id === tp.id)
      if (idx !== -1) topics.value[idx] = updated
      status.value = t('rules.saved', { topic: tp.topic })
    }
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  } finally {
    savingId.value = null
  }
}

/** Seed a candidate rule from the currently open email (subject + sender). */
function applyFromEmail(tp: RoutingTopic): void {
  const subjectWord = (item.value.subject || '').split(/\s+/).find((w) => w.length > 3) ?? ''
  const domain = (item.value.from ?? '').split('@')[1] ?? ''
  const parts: string[] = []
  if (subjectWord) parts.push(`subject contains "${subjectWord}"`)
  if (domain) parts.push(`from contains "@${domain}"`)
  const candidate = parts.length ? `IF ${parts.join(' OR ')} THEN topic=${tp.topic}` : ''
  if (candidate) {
    editRules.value[tp.id] = editRules.value[tp.id]
      ? `${editRules.value[tp.id]}\n${candidate}`
      : candidate
  }
}

async function testRouting(): Promise<void> {
  const text = `${item.value.subject}\n\n${item.value.bodyText}`.trim()
  if (!text) {
    error.value = t('rules.testNoEmail')
    return
  }
  testing.value = true
  error.value = null
  candidates.value = null
  try {
    const r = await call((c) => c.testRouting(text, 5))
    if (r) candidates.value = r.candidates
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  } finally {
    testing.value = false
  }
}

onMounted(() => void load())
</script>

<template>
  <section class="rule">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('settings')">
        ← {{ t('common.back') }}
      </button>
      <h2 class="syn-view-title">{{ t('settings.routingRules') }}</h2>
    </header>

    <p class="syn-muted">{{ t('rules.intro') }}</p>

    <div class="syn-card">
      <h3 class="syn-card-title">{{ t('rules.testTitle') }}</h3>
      <ActionButton :loading="testing" @click="testRouting">{{ t('rules.test') }}</ActionButton>
      <ul v-if="candidates && candidates.length" class="rule__candidates">
        <li v-for="(c, i) in candidates" :key="i">
          <strong>{{ c.topic }}</strong>
          <span class="rule__score">{{ c.score.toFixed(3) }}</span>
          <span v-if="c.stale" class="rule__stale">stale</span>
        </li>
      </ul>
      <p v-else-if="candidates && !candidates.length" class="syn-muted">
        {{ t('rules.testEmpty') }}
      </p>
    </div>

    <p v-if="loading" class="syn-muted">{{ t('rules.loading') }}</p>
    <p v-else-if="!topics.length" class="syn-muted">{{ t('rules.empty') }}</p>

    <div v-for="tp in topics" :key="tp.id" class="syn-card">
      <h3 class="syn-card-title">
        {{ tp.topic }}
        <span v-if="tp.isDefault" class="rule__badge">{{ t('rules.systemBadge') }}</span>
      </h3>
      <p class="syn-card-sub">{{ tp.shortDescription }}</p>

      <label class="rule__label">{{ t('rules.keywords') }}</label>
      <input
        v-model="editKeywords[tp.id]"
        type="text"
        class="rule__input"
        :disabled="tp.isDefault"
        :placeholder="t('rules.keywordsHint')"
      />

      <label class="rule__label">{{ t('rules.selectionRules') }}</label>
      <textarea v-model="editRules[tp.id]" rows="2" class="rule__input" :disabled="tp.isDefault" />

      <p v-if="tp.isDefault" class="syn-muted rule__readonly">{{ t('rules.readonly') }}</p>
      <div v-else class="syn-row">
        <ActionButton :block="false" :loading="savingId === tp.id" @click="saveTopic(tp)">
          {{ t('rules.save') }}
        </ActionButton>
        <ActionButton :block="false" @click="applyFromEmail(tp)">
          {{ t('rules.applyFromEmail') }}
        </ActionButton>
      </div>
    </div>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </section>
</template>

<style scoped>
.rule {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.rule__label {
  display: block;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
  margin-top: var(--syn-space-2);
}
.rule__input {
  width: 100%;
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  resize: vertical;
  box-sizing: border-box;
}
.rule__readonly {
  margin: var(--syn-space-2) 0 0;
  font-size: var(--syn-font-size-sm);
}
.rule__badge {
  font-size: var(--syn-font-size-sm);
  font-weight: 600;
  color: var(--syn-muted);
  border: 1px solid var(--syn-border);
  border-radius: 999px;
  padding: 1px var(--syn-space-2);
  margin-left: var(--syn-space-2);
}
.rule__candidates {
  list-style: none;
  padding: 0;
  margin: var(--syn-space-2) 0 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.rule__candidates li {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
  font-size: var(--syn-font-size-sm);
}
.rule__score {
  color: var(--syn-muted);
}
.rule__stale {
  color: var(--syn-danger, #b00);
  font-size: var(--syn-font-size-sm);
}
</style>
