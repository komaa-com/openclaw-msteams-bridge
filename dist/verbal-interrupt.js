const INTERRUPT_PHRASES = new Set([
    "stop",
    "stop it",
    "stop talking",
    "wait",
    "wait wait",
    "wait a second",
    "wait a minute",
    "hold on",
    "hang on",
    "never mind",
    "nevermind",
    "be quiet",
    "quiet",
    "shut up",
    "enough",
    "thats enough",
    "pause",
    "one second",
    "one sec",
    "give me a second",
    "توقف",
    "قف",
    "اسكت",
    "اصمت",
    "انتظر",
    "انتظر لحظة",
    "استنى",
    "استنا",
    "لحظة",
    "لحظه",
    "لحظة واحدة",
    "ثانية",
    "ثانيه",
    "ثانية واحدة",
    "دقيقة",
    "دقيقه",
    "مهلا",
    "خلاص",
    "كفى",
    "كفاية",
    "كفايه",
    "بس",
]);
const FILLER_TOKENS = new Set(["ok", "okay", "oh", "no", "hey", "please", "now", "طيب", "لا", "يا"]);
function normalizeWords(text) {
    return ((text ?? "")
        .toLowerCase()
        .replace(/['’]/g, "")
        .replace(/[\p{M}ـ]/gu, "")
        .replace(/[^\p{L}\s]/gu, " ")
        .split(/\s+/)
        .filter(Boolean));
}
function startsWithSeq(words, seq) {
    return seq.length > 0 && seq.length <= words.length && seq.every((w, i) => words[i] === w);
}
function endsWithSeq(words, seq) {
    const offset = words.length - seq.length;
    return seq.length > 0 && offset >= 0 && seq.every((w, i) => words[offset + i] === w);
}
export function isVerbalInterrupt(text, wakePhrases) {
    let core = normalizeWords(text);
    const wake = (wakePhrases ?? []).map(normalizeWords).filter((seq) => seq.length > 0);
    let changed = true;
    while (changed && core.length > 0) {
        changed = false;
        while (core.length > 0 && FILLER_TOKENS.has(core[0] ?? "")) {
            core.shift();
            changed = true;
        }
        while (core.length > 0 && FILLER_TOKENS.has(core[core.length - 1] ?? "")) {
            core.pop();
            changed = true;
        }
        for (const seq of wake) {
            if (startsWithSeq(core, seq)) {
                core = core.slice(seq.length);
                changed = true;
            }
            else if (endsWithSeq(core, seq)) {
                core = core.slice(0, core.length - seq.length);
                changed = true;
            }
        }
    }
    if (core.length === 0 || core.length > 4) {
        return false;
    }
    return INTERRUPT_PHRASES.has(core.join(" "));
}
