import type { SetNotesOp } from '../../shared/ipc'

export type SpeakerNotesWriteStatus = 'applied' | 'unchanged' | 'uncertain'

interface SpeakerNotesApi {
  setNotes(op: SetNotesOp): Promise<boolean>
  getNotes(slideIndex: number): Promise<string>
}

export async function setAndVerifySpeakerNotes(
  api: SpeakerNotesApi,
  slideIndex: number,
  text: string,
): Promise<SpeakerNotesWriteStatus> {
  if (!(await api.setNotes({ slideIndex, text }))) return 'unchanged'
  return (await api.getNotes(slideIndex)) === text ? 'applied' : 'uncertain'
}
