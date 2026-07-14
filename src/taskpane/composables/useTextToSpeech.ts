/**
 * Text-to-speech playback for chat answers. Fetches audio from Synaplan's
 * `GET /api/v1/tts/stream` (which uses the user's configured TEXT2SOUND model)
 * and plays it. Long answers are split into sentence-sized chunks so playback
 * starts quickly and stays within GET URL limits; chunks play back-to-back.
 *
 * Exactly one item plays at a time. `speakingId`/`loadingId` let a caller show
 * a speaker icon that flips to a stop button for the active item.
 */

import { onBeforeUnmount, ref } from 'vue'
import { standardLanguage } from './useLanguagePrefs'
import { useSynaplanClient } from './useSynaplanClient'

/** Max characters per TTS request — keeps GET URLs short and playback snappy. */
const MAX_CHUNK = 480

/** Split text into <= MAX_CHUNK pieces, preferring sentence boundaries. */
export function splitForTts(text: string): string[] {
  const clean = text.trim()
  if (clean.length === 0) return []
  if (clean.length <= MAX_CHUNK) return [clean]

  const sentences = clean.split(/(?<=[.!?…])\s+/)
  const out: string[] = []
  let current = ''
  for (const sentence of sentences) {
    if (sentence.length > MAX_CHUNK) {
      if (current) {
        out.push(current)
        current = ''
      }
      for (let i = 0; i < sentence.length; i += MAX_CHUNK) {
        out.push(sentence.slice(i, i + MAX_CHUNK))
      }
      continue
    }
    const candidate = current ? `${current} ${sentence}` : sentence
    if (candidate.length > MAX_CHUNK) {
      if (current) out.push(current)
      current = sentence
    } else {
      current = candidate
    }
  }
  if (current) out.push(current)
  return out
}

export function useTextToSpeech() {
  const { call } = useSynaplanClient()

  /** Id of the item currently playing audio, or null. */
  const speakingId = ref<string | null>(null)
  /** Id of the item whose first chunk is being fetched, or null. */
  const loadingId = ref<string | null>(null)

  let audio: HTMLAudioElement | null = null
  let objectUrl: string | null = null
  // Bumped on every stop()/toggle() so an in-flight fetch/playback sequence can
  // detect it has been superseded and bail out.
  let runToken = 0

  function releaseAudio(): void {
    if (audio) {
      audio.onended = null
      audio.onerror = null
      audio.pause()
      audio.src = ''
      audio = null
    }
    if (objectUrl) {
      URL.revokeObjectURL(objectUrl)
      objectUrl = null
    }
  }

  function stop(): void {
    runToken++
    releaseAudio()
    speakingId.value = null
    loadingId.value = null
  }

  function playBlob(blob: Blob): Promise<void> {
    return new Promise((resolve, reject) => {
      releaseAudio()
      objectUrl = URL.createObjectURL(blob)
      const el = new Audio(objectUrl)
      audio = el
      el.onended = () => resolve()
      el.onerror = () => reject(new Error('TTS audio playback failed'))
      void el.play().catch(reject)
    })
  }

  /**
   * Toggle playback for `id`: start reading `text` aloud, or stop if that id is
   * already playing/loading. Starting cancels any other active playback.
   */
  async function toggle(id: string, text: string, language?: string): Promise<void> {
    if (speakingId.value === id || loadingId.value === id) {
      stop()
      return
    }
    stop()
    const myToken = runToken
    const chunks = splitForTts(text)
    if (chunks.length === 0) return

    const lang = language ?? standardLanguage()
    loadingId.value = id
    try {
      for (const part of chunks) {
        if (runToken !== myToken) return
        const blob = await call((c) => c.tts({ text: part, language: lang }))
        if (runToken !== myToken || !blob) return
        loadingId.value = null
        speakingId.value = id
        await playBlob(blob)
        if (runToken !== myToken) return
      }
    } finally {
      if (runToken === myToken) {
        speakingId.value = null
        loadingId.value = null
      }
    }
  }

  onBeforeUnmount(stop)

  return { speakingId, loadingId, toggle, stop }
}
