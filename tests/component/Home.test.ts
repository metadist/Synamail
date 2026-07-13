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
    ragGroups: vi.fn(async () => []),
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
  it('renders the Ask-synaplan chat composer', () => {
    const wrapper = mountHome()
    expect(wrapper.find('textarea').exists()).toBe(true)
    expect(wrapper.text()).toContain(en.home.commands.chat)
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

  it('renders the four function boxes', () => {
    const text = mountHome().text()
    expect(text).toContain(en.home.boxes.emailWriting.title)
    expect(text).toContain(en.home.boxes.summarize.title)
    expect(text).toContain(en.home.boxes.knowledge.title)
    expect(text).toContain(en.home.commands.chat)
  })

  it('offers the three writing styles in the email box', () => {
    const text = mountHome().text()
    expect(text).toContain(en.tone.concise)
    expect(text).toContain(en.tone.detailed)
    expect(text).toContain(en.tone.formal)
  })

  it('hides the (temporarily disabled) profiling section', () => {
    const text = mountHome().text()
    expect(text).not.toContain(en.home.sections.profiling)
  })
})
