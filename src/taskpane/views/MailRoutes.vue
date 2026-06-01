<script setup lang="ts">
import { onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { getMailRoutes, saveMailRoutes } from '@/taskpane/composables/useRoamingSettings'
import { go } from '@/taskpane/router'
import { SUPPORTED_LOCALES } from '@/i18n'
import type { MailRoute, MailRouteKind } from '@shared/mail-routes/types'

const { t } = useI18n()

const routes = ref<MailRoute[]>([])
const paused = ref(false)
const seen = ref<Record<string, number>>({})
const error = ref<string | null>(null)
const status = ref<string | null>(null)

// Editor state — a flat set of draft fields, assembled into a typed MailRoute
// on save so the template never has to narrow the discriminated union.
const isEditing = ref(false)
const draftId = ref('')
const draftKind = ref<MailRouteKind>('meeting')
const draftName = ref('')
const draftEnabled = ref(true)
const draftSenders = ref('')
const draftDuration = ref(30)
const draftGroupId = ref('')
const draftTag = ref('')
const draftWorkingLanguage = ref<string>('en')
const draftMinBodyChars = ref(200)
const draftIncludeAttachments = ref(true)

const KINDS: MailRouteKind[] = ['meeting', 'projectIngest', 'newsletterKb']

function load(): void {
  const state = getMailRoutes()
  routes.value = state.routes
  paused.value = state.paused
  seen.value = state.seen
}

async function persist(): Promise<void> {
  error.value = null
  try {
    await saveMailRoutes({ paused: paused.value, routes: routes.value, seen: seen.value })
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
  }
}

function genId(): string {
  const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto
  if (c?.randomUUID) return c.randomUUID()
  return `route-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function parseSenders(text: string): string[] {
  const seenSet = new Set<string>()
  const out: string[] = []
  for (const raw of text.split(/[\s,;]+/)) {
    const v = raw.trim().toLowerCase()
    if (v && !seenSet.has(v)) {
      seenSet.add(v)
      out.push(v)
    }
  }
  return out
}

function kindLabel(kind: MailRouteKind): string {
  return t(`mailRoutes.kind.${kind}`)
}

function routeSummary(route: MailRoute): string {
  if (route.kind === 'meeting') {
    return t('mailRoutes.summary.meeting', { mins: route.durationMinutes })
  }
  return t('mailRoutes.summary.ingest', { group: route.groupId, lang: route.workingLanguage })
}

function beginCreate(kind: MailRouteKind): void {
  draftId.value = ''
  draftKind.value = kind
  draftName.value = ''
  draftEnabled.value = true
  draftSenders.value = ''
  draftDuration.value = 30
  draftGroupId.value = kind === 'newsletterKb' ? 'newsletters' : ''
  draftTag.value = ''
  draftWorkingLanguage.value = 'en'
  draftMinBodyChars.value = kind === 'newsletterKb' ? 400 : 200
  draftIncludeAttachments.value = kind === 'projectIngest'
  isEditing.value = true
}

function beginEdit(route: MailRoute): void {
  draftId.value = route.id
  draftKind.value = route.kind
  draftName.value = route.name
  draftEnabled.value = route.enabled
  draftSenders.value = route.senders.join(', ')
  if (route.kind === 'meeting') {
    draftDuration.value = route.durationMinutes
  } else {
    draftGroupId.value = route.groupId
    draftTag.value = route.tag
    draftWorkingLanguage.value = route.workingLanguage
    draftMinBodyChars.value = route.minBodyChars
    draftIncludeAttachments.value = route.includeAttachments
  }
  isEditing.value = true
}

function cancelEdit(): void {
  isEditing.value = false
}

function buildDraft(): MailRoute | null {
  const senders = parseSenders(draftSenders.value)
  const name = draftName.value.trim()
  if (!name || senders.length === 0) {
    error.value = t('mailRoutes.validation.nameAndSenders')
    return null
  }
  const id = draftId.value || genId()
  if (draftKind.value === 'meeting') {
    return {
      id,
      kind: 'meeting',
      name,
      enabled: draftEnabled.value,
      senders,
      durationMinutes: Math.max(5, Math.round(draftDuration.value) || 30),
    }
  }
  const groupId = draftGroupId.value.trim()
  if (!groupId) {
    error.value = t('mailRoutes.validation.group')
    return null
  }
  return {
    id,
    kind: draftKind.value,
    name,
    enabled: draftEnabled.value,
    senders,
    groupId,
    tag: draftTag.value.trim() || groupId,
    workingLanguage: draftWorkingLanguage.value,
    minBodyChars: Math.max(0, Math.round(draftMinBodyChars.value) || 0),
    includeAttachments: draftIncludeAttachments.value,
  }
}

async function saveDraft(): Promise<void> {
  error.value = null
  const route = buildDraft()
  if (!route) return
  const idx = routes.value.findIndex((r) => r.id === route.id)
  if (idx === -1) routes.value.push(route)
  else routes.value[idx] = route
  await persist()
  status.value = t('mailRoutes.saved', { name: route.name })
  isEditing.value = false
}

async function removeRoute(route: MailRoute): Promise<void> {
  routes.value = routes.value.filter((r) => r.id !== route.id)
  await persist()
  status.value = t('mailRoutes.removed', { name: route.name })
}

async function toggleEnabled(route: MailRoute): Promise<void> {
  route.enabled = !route.enabled
  await persist()
}

async function togglePaused(): Promise<void> {
  paused.value = !paused.value
  await persist()
}

onMounted(load)
</script>

<template>
  <section class="routes">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('settings')">
        ← {{ t('common.back') }}
      </button>
      <h2 class="syn-view-title">{{ t('mailRoutes.title') }}</h2>
    </header>

    <p class="syn-muted">{{ t('mailRoutes.intro') }}</p>

    <div class="syn-card routes__pause">
      <div>
        <h3 class="syn-card-title">{{ t('mailRoutes.pauseTitle') }}</h3>
        <p class="syn-card-sub">
          {{ paused ? t('mailRoutes.pausedOn') : t('mailRoutes.pausedOff') }}
        </p>
      </div>
      <ActionButton :block="false" @click="togglePaused">
        {{ paused ? t('mailRoutes.resume') : t('mailRoutes.pause') }}
      </ActionButton>
    </div>

    <!-- Editor -->
    <div v-if="isEditing" class="syn-card">
      <h3 class="syn-card-title">
        {{ draftId ? t('mailRoutes.editTitle') : t('mailRoutes.newTitle') }} —
        {{ kindLabel(draftKind) }}
      </h3>

      <label class="routes__label">{{ t('mailRoutes.fields.name') }}</label>
      <input
        v-model="draftName"
        type="text"
        :placeholder="t('mailRoutes.fields.namePlaceholder')"
      />

      <label class="routes__label">{{ t('mailRoutes.fields.senders') }}</label>
      <textarea
        v-model="draftSenders"
        rows="2"
        :placeholder="t('mailRoutes.fields.sendersPlaceholder')"
      />
      <p class="syn-muted routes__hint">{{ t('mailRoutes.fields.sendersHint') }}</p>

      <template v-if="draftKind === 'meeting'">
        <label class="routes__label">{{ t('mailRoutes.fields.duration') }}</label>
        <input v-model.number="draftDuration" type="number" min="5" step="5" />
      </template>

      <template v-else>
        <label class="routes__label">{{ t('mailRoutes.fields.group') }}</label>
        <input
          v-model="draftGroupId"
          type="text"
          :placeholder="t('mailRoutes.fields.groupPlaceholder')"
        />

        <label class="routes__label">{{ t('mailRoutes.fields.tag') }}</label>
        <input
          v-model="draftTag"
          type="text"
          :placeholder="t('mailRoutes.fields.tagPlaceholder')"
        />

        <label class="routes__label">{{ t('mailRoutes.fields.workingLanguage') }}</label>
        <select v-model="draftWorkingLanguage" class="routes__select">
          <option v-for="l in SUPPORTED_LOCALES" :key="l" :value="l">
            {{ t(`language.${l}`) }}
          </option>
        </select>
        <p class="syn-muted routes__hint">{{ t('mailRoutes.fields.workingLanguageHint') }}</p>

        <label class="routes__label">{{ t('mailRoutes.fields.minBodyChars') }}</label>
        <input v-model.number="draftMinBodyChars" type="number" min="0" step="50" />

        <label class="routes__check">
          <input v-model="draftIncludeAttachments" type="checkbox" />
          <span>{{ t('mailRoutes.fields.includeAttachments') }}</span>
        </label>
      </template>

      <label class="routes__check">
        <input v-model="draftEnabled" type="checkbox" />
        <span>{{ t('mailRoutes.fields.enabled') }}</span>
      </label>

      <div class="syn-row">
        <ActionButton :block="false" variant="primary" @click="saveDraft">
          {{ t('common.save') }}
        </ActionButton>
        <ActionButton :block="false" @click="cancelEdit">{{ t('common.cancel') }}</ActionButton>
      </div>
    </div>

    <!-- List -->
    <template v-else>
      <p v-if="!routes.length" class="syn-muted">{{ t('mailRoutes.empty') }}</p>

      <div v-for="r in routes" :key="r.id" class="syn-card" :class="{ routes__off: !r.enabled }">
        <h3 class="syn-card-title">
          {{ r.name }}
          <span class="routes__badge">{{ kindLabel(r.kind) }}</span>
        </h3>
        <p class="syn-card-sub">{{ r.senders.join(', ') }}</p>
        <p class="syn-card-sub">{{ routeSummary(r) }}</p>
        <div class="syn-row">
          <ActionButton :block="false" @click="toggleEnabled(r)">
            {{ r.enabled ? t('mailRoutes.disable') : t('mailRoutes.enable') }}
          </ActionButton>
          <ActionButton :block="false" @click="beginEdit(r)">{{
            t('mailRoutes.edit')
          }}</ActionButton>
          <ActionButton :block="false" @click="removeRoute(r)">
            {{ t('mailRoutes.remove') }}
          </ActionButton>
        </div>
      </div>

      <div class="syn-card">
        <h3 class="syn-card-title">{{ t('mailRoutes.addTitle') }}</h3>
        <p class="syn-card-sub">{{ t('mailRoutes.addHint') }}</p>
        <div class="routes__add">
          <ActionButton v-for="k in KINDS" :key="k" :block="false" @click="beginCreate(k)">
            + {{ kindLabel(k) }}
          </ActionButton>
        </div>
      </div>
    </template>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />
  </section>
</template>

<style scoped>
.routes {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.routes__pause {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--syn-space-2);
}
.routes__label {
  display: block;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
  margin-top: var(--syn-space-2);
}
.routes__hint {
  margin: var(--syn-space-1) 0 0;
  font-size: var(--syn-font-size-sm);
}
.routes input[type='text'],
.routes input[type='number'],
.routes textarea,
.routes__select {
  width: 100%;
  box-sizing: border-box;
}
.routes__select {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
}
.routes__check {
  display: flex;
  align-items: center;
  gap: var(--syn-space-2);
  margin-top: var(--syn-space-2);
  font-size: var(--syn-font-size-sm);
}
.routes__check input {
  width: auto;
}
.routes__badge {
  font-size: var(--syn-font-size-sm);
  font-weight: 600;
  color: var(--syn-muted);
  border: 1px solid var(--syn-border);
  border-radius: 999px;
  padding: 1px var(--syn-space-2);
  margin-left: var(--syn-space-2);
}
.routes__off {
  opacity: 0.6;
}
.routes__add {
  display: flex;
  flex-wrap: wrap;
  gap: var(--syn-space-2);
  margin-top: var(--syn-space-2);
}
</style>
