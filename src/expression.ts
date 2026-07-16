/**
 * Avatar expression cues.
 *
 * Infers a coarse emotion tag from the assistant's reply text and sends it on the wire as a cue for
 * a receiver's avatar. Deliberately a cheap, dependency-free lexical heuristic: it runs on every turn
 * with no extra model call or latency, is cosmetic only (audio is untouched), and a "neutral" result
 * is always safe. Best-effort: a receiver that doesn't recognise the tag ignores it.
 */

export type AvatarEmotion = "neutral" | "happy" | "sad" | "surprised";

/** "!!", "?!", "!?" — any run of 2+ emphatic marks reads as surprise regardless of wording. */
const SURPRISED_PUNCT = /[?!]{2,}/;
const SURPRISED_WORDS =
  /\b(wow|whoa|woah|oh no|oh my|no way|unbelievable|incredible|astonish\w*|surpris\w*)\b/;
const SAD_WORDS =
  /\b(sorry|apolog\w*|unfortunately|regret\w*|afraid|sadly|bad news|failed|unable to|i can'?t|i cannot|i'?m unable)\b/;
const HAPPY_WORDS =
  /\b(glad|great|awesome|wonderful|fantastic|excellent|congrat\w*|happy|love|perfect|good news|success\w*|thank\w*|welcome|nice|well done)\b/;

/**
 * Map reply text to an {@link AvatarEmotion}. Surprise wins over sad over happy (a strong "wow!"
 * shouldn't be masked by a polite "thanks"); empty/neutral text → "neutral".
 *
 * Known limitation: this is first-match priority, not sentiment scoring — a mixed reply like
 * "Unfortunately I'm glad I could help!" resolves to "sad" (the higher-priority match wins). That's an
 * acceptable trade-off for a cheap cosmetic cue; the realtime path also re-cues as more text arrives.
 */
export function inferEmotion(text: string | null | undefined): AvatarEmotion {
  const raw = (text ?? "").trim();
  if (!raw) {
    return "neutral";
  }
  const lower = raw.toLowerCase();
  if (SURPRISED_PUNCT.test(raw) || SURPRISED_WORDS.test(lower)) {
    return "surprised";
  }
  if (SAD_WORDS.test(lower)) {
    return "sad";
  }
  if (HAPPY_WORDS.test(lower)) {
    return "happy";
  }
  return "neutral";
}
