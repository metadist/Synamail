import { describe, expect, it, vi } from 'vitest'
import { flushPromises, mount } from '@vue/test-utils'
import { createI18n } from 'vue-i18n'
import en from '@/locales/en.json'

// Hoisted spies so the vi.mock factories (hoisted above imports) can reference them.
const { composeDraft, setComposeBody, itemRef } = vi.hoisted(() => ({
  composeDraft: vi.fn(async () => ({ htmlBody: '<p>Generated body</p>' })),
  setComposeBody: vi.fn(async () => true),
  itemRef: {
    value: {
      mode: 'compose' as const,
      subject: '',
      from: undefined,
      to: [] as string[],
      cc: [] as string[],
      bodyText: '',
      attachments: [] as unknown[],
    },
  },
}))

vi.mock('@/taskpane/composables/useSynaplanClient', () => {
  const fakeClient = { composeDraft }
  return {
    AUTH_INVALIDATED_EVENT: 'synamail:auth-invalidated',
    useSynaplanClient: () => ({
      call: async <T>(fn: (c: typeof fakeClient) => Promise<T>) => fn(fakeClient),
      client: { value: fakeClient },
      baseUrl: { value: 'https://x' },
    }),
  }
})

vi.mock('@/taskpane/composables/useOutlookItem', () => ({
  useOutlookItem: () => ({ item: itemRef }),
  setComposeBody,
}))

import ComposeAssistantPanel from '@/taskpane/components/ComposeAssistantPanel.vue'

const i18n = createI18n({ legacy: false, locale: 'en', messages: { en } })

function mountPanel() {
  return mount(ComposeAssistantPanel, { global: { plugins: [i18n] } })
}

describe('ComposeAssistantPanel.vue', () => {
  it('renders the intent form when composing an email', () => {
    itemRef.value.mode = 'compose'
    const wrapper = mountPanel()
    expect(wrapper.find('textarea').exists()).toBe(true)
    expect(wrapper.text()).not.toContain(en.compose.noCompose)
  })

  it('shows a hint (no form) when not in compose mode', () => {
    itemRef.value.mode = 'read' as unknown as 'compose'
    const wrapper = mountPanel()
    expect(wrapper.find('textarea').exists()).toBe(false)
    expect(wrapper.text()).toContain(en.compose.noCompose)
    itemRef.value.mode = 'compose'
  })

  it('drafts from the intent and writes the HTML into the compose body', async () => {
    itemRef.value.mode = 'compose'
    const wrapper = mountPanel()
    await wrapper.find('textarea').setValue('Invite Alice to lunch on Friday')
    await wrapper.find('button.ab--primary').trigger('click')
    await flushPromises()
    expect(composeDraft).toHaveBeenCalledWith(
      expect.objectContaining({
        intent: 'Invite Alice to lunch on Friday',
        tone: 'concise',
        language: 'en',
      }),
    )
    expect(setComposeBody).toHaveBeenCalledWith('<p>Generated body</p>')
    expect(wrapper.text()).toContain(en.compose.inserted)
  })
})
