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

  it('always shows the Synaplan server URL field so the user can pick the target', () => {
    const wrapper = mount(SignIn, { global: { plugins: [i18n] } })
    expect(wrapper.find('input[type=url]').exists()).toBe(true)
  })

  it('shows a Save button only once the server URL changes', async () => {
    const wrapper = mount(SignIn, { global: { plugins: [i18n] } })
    expect(wrapper.findAll('button').some((b) => b.text() === 'Save')).toBe(false)
    await wrapper.find('input[type=url]').setValue('https://my.synaplan.example')
    expect(wrapper.findAll('button').some((b) => b.text() === 'Save')).toBe(true)
  })

  it('rejects an http instance with an HTTPS hint instead of signing in', async () => {
    const wrapper = mount(SignIn, { global: { plugins: [i18n] } })
    await wrapper.find('input[type=url]').setValue('http://localhost/')
    await wrapper.find('button.ab--primary').trigger('click')
    expect(wrapper.find('.toast--error').text()).toMatch(/HTTPS/i)
  })
})
