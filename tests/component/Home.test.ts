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
  it('renders the chat composer and the two command launchers', () => {
    const wrapper = mountHome()
    expect(wrapper.find('textarea').exists()).toBe(true)
    expect(wrapper.findAll('button.home__cmd')).toHaveLength(2)
  })

  it('sends a chat message and shows the AI reply', async () => {
    const wrapper = mountHome()
    await wrapper.find('textarea').setValue('hello world')
    await wrapper.find('button.ab--primary').trigger('click')
    await flushPromises()
    expect(wrapper.text()).toContain('hello world')
    expect(wrapper.text()).toContain('echo: hello world')
  })

  it('opens the search & vectorize dialog from the first command', async () => {
    const wrapper = mountHome()
    expect(wrapper.find('#sv-topic').exists()).toBe(false)
    await wrapper.findAll('button.home__cmd')[0].trigger('click')
    expect(wrapper.find('#sv-topic').exists()).toBe(true)
  })

  it('opens the new mail dialog from the second command', async () => {
    const wrapper = mountHome()
    expect(wrapper.find('#nm-desc').exists()).toBe(false)
    await wrapper.findAll('button.home__cmd')[1].trigger('click')
    expect(wrapper.find('#nm-desc').exists()).toBe(true)
  })
})
