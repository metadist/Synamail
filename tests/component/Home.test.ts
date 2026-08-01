import { describe, expect, it, vi } from 'vitest'
import { flushPromises, mount } from '@vue/test-utils'
import { createI18n } from 'vue-i18n'
import en from '@/locales/en.json'

const { fakeClient, itemRef } = vi.hoisted(() => {
  const itemRef = {
    value: {
      mode: 'none' as 'none' | 'read' | 'compose',
      subject: '',
      to: [] as string[],
      cc: [] as string[],
      bodyText: '',
      attachments: [] as unknown[],
      conversationId: undefined as string | undefined,
    },
  }
  const fakeClient = {
    chat: vi.fn(async (input: { chatId?: number; question: string; emailContext?: string }) => ({
      chatId: input.chatId ?? 5,
      answer: `echo: ${input.question}`,
    })),
    getChatMessages: vi.fn(async () => [] as { role: 'user' | 'ai'; text: string }[]),
    ragGroups: vi.fn(async () => []),
  }
  return { fakeClient, itemRef }
})

vi.mock('@/taskpane/composables/useSynaplanClient', () => {
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
  useOutlookItem: () => ({
    item: itemRef,
    loading: { value: false },
    error: { value: null },
    refresh: vi.fn(),
  }),
}))

import Home from '@/taskpane/views/Home.vue'

const i18n = createI18n({ legacy: false, locale: 'en', messages: { en } })

function resetItem(): void {
  itemRef.value = {
    mode: 'none',
    subject: '',
    to: [],
    cc: [],
    bodyText: '',
    attachments: [],
    conversationId: undefined,
  }
}

function seedRoamingChat(conversationId: string, chatId: number): void {
  Office.context.roamingSettings.set('synamail.settings', {
    apiKey: 'k',
    keyId: 1,
    email: 'a@b.test',
    baseUrl: 'https://x',
    chats: { [conversationId]: chatId },
  })
}

function mountHome() {
  return mount(Home, { global: { plugins: [i18n] } })
}

/** Chat Send — not Knowledge Base "New", which is also `ab--primary` when mail is open. */
function sendButton(wrapper: ReturnType<typeof mountHome>) {
  return wrapper.findAll('button.ab--primary').find((b) => b.text() === en.home.chat.send)!
}

describe('Home.vue', () => {
  it('renders the Ask-synaplan chat composer', () => {
    resetItem()
    const wrapper = mountHome()
    expect(wrapper.find('textarea').exists()).toBe(true)
    expect(wrapper.text()).toContain(en.home.commands.chat)
  })

  it('starts with an empty composer and Send disabled', () => {
    resetItem()
    const wrapper = mountHome()
    const textarea = wrapper.find('textarea').element as HTMLTextAreaElement
    expect(textarea.value).toBe('')
    expect(sendButton(wrapper).attributes('disabled')).toBeDefined()
  })

  it('sends a chat message and shows the AI reply', async () => {
    resetItem()
    fakeClient.chat.mockClear()
    const wrapper = mountHome()
    await wrapper.find('textarea').setValue('hello world')
    await sendButton(wrapper).trigger('click')
    await flushPromises()
    expect(wrapper.text()).toContain('hello world')
    expect(wrapper.text()).toContain('echo: hello world')
  })

  it('grounds Home chat on the open email body', async () => {
    fakeClient.chat.mockClear()
    itemRef.value = {
      mode: 'read',
      subject: 'Invoice',
      to: [],
      cc: [],
      bodyText: 'Please approve invoice 42.',
      attachments: [],
      conversationId: 'conv-1',
    }
    const wrapper = mountHome()
    await flushPromises()
    await wrapper.find('textarea').setValue('what is this about?')
    await sendButton(wrapper).trigger('click')
    await flushPromises()
    expect(fakeClient.chat).toHaveBeenCalled()
    const arg = fakeClient.chat.mock.calls[0][0] as {
      emailContext?: string
      conversationId: string
      mailSubject?: string
    }
    expect(arg.emailContext).toBe('Please approve invoice 42.')
    expect(arg.conversationId).toBe('conv-1')
    expect(arg.mailSubject).toBe('Invoice')
    resetItem()
  })

  it('restores prior turns from Synaplan when a roaming chat id exists', async () => {
    fakeClient.getChatMessages.mockClear()
    fakeClient.getChatMessages.mockResolvedValueOnce([
      { role: 'user', text: 'earlier question' },
      { role: 'ai', text: 'earlier answer' },
    ])
    seedRoamingChat('conv-restore', 99)
    itemRef.value = {
      mode: 'read',
      subject: 'Thread',
      to: [],
      cc: [],
      bodyText: 'body',
      attachments: [],
      conversationId: 'conv-restore',
    }
    const wrapper = mountHome()
    await flushPromises()
    expect(fakeClient.getChatMessages).toHaveBeenCalledWith(99)
    expect(wrapper.text()).toContain('earlier question')
    expect(wrapper.text()).toContain('earlier answer')
    resetItem()
  })

  it('renders the four function boxes', () => {
    resetItem()
    const text = mountHome().text()
    expect(text).toContain(en.home.boxes.emailWriting.title)
    expect(text).toContain(en.home.boxes.summarize.title)
    expect(text).toContain(en.home.boxes.knowledge.title)
    expect(text).toContain(en.home.commands.chat)
  })

  it('offers the three writing styles in the email box', () => {
    resetItem()
    const text = mountHome().text()
    expect(text).toContain(en.tone.concise)
    expect(text).toContain(en.tone.detailed)
    expect(text).toContain(en.tone.formal)
  })

  it('hides the (temporarily disabled) profiling section', () => {
    resetItem()
    const text = mountHome().text()
    expect(text).not.toContain(en.home.sections.profiling)
  })
})
