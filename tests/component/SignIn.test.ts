import { describe, expect, it } from 'vitest'
import { mount } from '@vue/test-utils'
import { createI18n } from 'vue-i18n'
import SignIn from '@/taskpane/views/SignIn.vue'
import en from '@/locales/en.json'

const i18n = createI18n({ legacy: false, locale: 'en', messages: { en } })

describe('SignIn.vue', () => {
  it('renders the primary sign-in button', () => {
    const wrapper = mount(SignIn, { global: { plugins: [i18n] } })
    expect(wrapper.text()).toContain('Sign in to Synaplan')
  })

  it('toggles the self-hosted override input', async () => {
    const wrapper = mount(SignIn, { global: { plugins: [i18n] } })
    expect(wrapper.find('input[type=url]').exists()).toBe(false)
    await wrapper.find('button.signin__link').trigger('click')
    expect(wrapper.find('input[type=url]').exists()).toBe(true)
  })
})
