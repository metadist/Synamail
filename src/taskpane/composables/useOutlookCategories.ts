/**
 * Outlook native categories — read the master list and apply/remove a category
 * on the current item. Used by the Mail Routes "categorize" action.
 *
 * Categories are a Mailbox 1.8 feature and writing them needs ReadWriteItem /
 * ReadWriteMailbox (we have ReadWriteMailbox). All calls feature-detect and
 * resolve gracefully when unsupported (older hosts, tests).
 */

interface MasterCategory {
  displayName: string
  color: Office.MailboxEnums.CategoryColor
}

function mailbox(): (Office.Mailbox & { masterCategories?: Office.MasterCategories }) | undefined {
  return typeof Office !== 'undefined' ? Office.context?.mailbox : undefined
}

export function categoriesSupported(): boolean {
  const mb = mailbox()
  const item = mb?.item as { categories?: unknown } | undefined
  return !!mb?.masterCategories && !!item && 'categories' in (item ?? {})
}

/** Best-effort preset color for a repurposed colour-name category. */
function presetFor(name: string): Office.MailboxEnums.CategoryColor {
  const C = Office.MailboxEnums.CategoryColor
  const n = name.toLowerCase()
  if (n.includes('red')) return C.Preset0
  if (n.includes('orange')) return C.Preset1
  if (n.includes('yellow')) return C.Preset5
  if (n.includes('green')) return C.Preset6
  if (n.includes('blue')) return C.Preset9
  if (n.includes('purple')) return C.Preset11
  return C.Preset0
}

export function getMasterCategories(): Promise<MasterCategory[]> {
  const mb = mailbox()
  if (!mb?.masterCategories) return Promise.resolve([])
  return new Promise((resolve) => {
    mb.masterCategories!.getAsync((r) => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value as MasterCategory[]) : [])
    })
  })
}

/** Ensure a category with this name exists in the mailbox master list. */
async function ensureMasterCategory(name: string): Promise<void> {
  const mb = mailbox()
  if (!mb?.masterCategories) return
  const existing = await getMasterCategories()
  if (existing.some((c) => c.displayName.toLowerCase() === name.toLowerCase())) return
  await new Promise<void>((resolve) => {
    mb.masterCategories!.addAsync([{ displayName: name, color: presetFor(name) }], () => resolve())
  })
}

/** Apply a category to the current read/compose item (creating it if needed). */
export async function applyCategory(name: string): Promise<void> {
  const mb = mailbox()
  const item = mb?.item as
    | {
        categories?: {
          addAsync?: (n: string[], cb: (r: Office.AsyncResult<void>) => void) => void
        }
      }
    | undefined
  if (!item?.categories?.addAsync) {
    throw new Error('This Outlook host does not support categories.')
  }
  await ensureMasterCategory(name)
  await new Promise<void>((resolve, reject) => {
    item.categories!.addAsync!([name], (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new Error(r.error?.message ?? 'Failed to apply category'))
    })
  })
}

/** Remove a category from the current item (used for undo). */
export async function removeCategory(name: string): Promise<void> {
  const mb = mailbox()
  const item = mb?.item as
    | {
        categories?: {
          removeAsync?: (n: string[], cb: (r: Office.AsyncResult<void>) => void) => void
        }
      }
    | undefined
  if (!item?.categories?.removeAsync) return
  await new Promise<void>((resolve) => {
    item.categories!.removeAsync!([name], () => resolve())
  })
}
