import { describe, expect, it, vi } from 'vitest'
import { flushPromises, mount } from '@vue/test-utils'
import { createI18n } from 'vue-i18n'
import en from '@/locales/en.json'

// Stub the Synaplan client so the chat round-trip is synchronous and offline.
vi.mock('@/taskpane/composables/useSynaplanClient', () => {
  const fakeClient = {
    chat: vi.fn(async (input: { chatId?: number; question: string }) => ({
      chatId: input.chatId ?? 5,
      answer: `echo: ${input.question}`,
    })),
  }
  return {
    AUTH_INVALIDATED_EVENT: 'synamail:auth-invalidated',
    useSynaplanClient: () => ({
      call: async <T>(fn: (c: typeof fakeClient) => Promise<T>) => fn(fakeClient),
      client: { value: fakeClient },
      baseUrl: { value: 'https://x' },
    }),
  }
})

import Home from '@/taskpane/views/Home.vue'

const i18n = createI18n({ legacy: false, locale: 'en', messages: { en } })

function mountHome() {
  return mount(Home, { global: { plugins: [i18n] } })
}

describe('Home.vue', () => {
  it('renders the chat composer at the top', () => {
    const wrapper = mountHome()
    expect(wrapper.find('textarea').exists()).toBe(true)
  })

  it('seeds the composer with a sample question so Send starts active', () => {
    const wrapper = mountHome()
    const textarea = wrapper.find('textarea').element as HTMLTextAreaElement
    expect(textarea.value).toBe(en.home.chat.sample)
    expect(wrapper.find('button.ab--primary').attributes('disabled')).toBeUndefined()
  })

  it('sends a chat message and shows the AI reply', async () => {
    const wrapper = mountHome()
    await wrapper.find('textarea').setValue('hello world')
    await wrapper.find('button.ab--primary').trigger('click')
    await flushPromises()
    expect(wrapper.text()).toContain('hello world')
    expect(wrapper.text()).toContain('echo: hello world')
  })

  it('renders the email-actions accordion section', () => {
    const text = mountHome().text()
    expect(text).toContain(en.home.sections.emailActions)
    // Cut features stay cut: no knowledge-filter or mail-routes sections.
    expect(text).not.toContain('Filter for knowledge base')
    expect(text).not.toContain('Email automations')
  })

  it('renders the writing-assistant accordion section', () => {
    const text = mountHome().text()
    expect(text).toContain(en.home.sections.writingAssistant)
  })
})
