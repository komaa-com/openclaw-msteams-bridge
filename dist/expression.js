const SURPRISED_PUNCT = /[?!]{2,}/;
const SURPRISED_WORDS = /\b(wow|whoa|woah|oh no|oh my|no way|unbelievable|incredible|astonish\w*|surpris\w*)\b/;
const SAD_WORDS = /\b(sorry|apolog\w*|unfortunately|regret\w*|afraid|sadly|bad news|failed|unable to|i can'?t|i cannot|i'?m unable)\b/;
const HAPPY_WORDS = /\b(glad|great|awesome|wonderful|fantastic|excellent|congrat\w*|happy|love|perfect|good news|success\w*|thank\w*|welcome|nice|well done)\b/;
export function inferEmotion(text) {
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
