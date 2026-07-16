const CHAR_VISEME = {
    a: 2,
    e: 4,
    i: 6,
    o: 8,
    u: 7,
    y: 6,
    m: 21,
    b: 21,
    p: 21,
    f: 18,
    v: 18,
    w: 7,
    r: 13,
    l: 14,
    s: 15,
    z: 15,
    x: 15,
    t: 19,
    d: 19,
    n: 19,
    k: 20,
    g: 20,
    c: 20,
    q: 20,
    h: 12,
    j: 16,
    ا: 2,
    أ: 2,
    إ: 2,
    آ: 2,
    ى: 2,
    ة: 2,
    و: 7,
    ي: 6,
    ئ: 6,
    م: 21,
    ب: 21,
    ف: 18,
    ر: 13,
    ل: 14,
    س: 15,
    ص: 15,
    ز: 15,
    ش: 16,
    ج: 16,
    ت: 19,
    د: 19,
    ن: 19,
    ط: 19,
    ض: 19,
    ث: 19,
    ذ: 19,
    ظ: 19,
    ك: 20,
    ق: 20,
    غ: 20,
    خ: 20,
    ه: 12,
    ح: 12,
    ع: 12,
    ء: 12,
    ؤ: 12,
    "َ": 2,
    "ُ": 7,
    "ِ": 6,
};
export function estimateVisemes(text, durationMs) {
    const normalized = (text ?? "").toLowerCase().replace(/\s+/g, " ").trim();
    if (!normalized || durationMs <= 0) {
        return [];
    }
    const tokens = [];
    for (const ch of normalized) {
        if (ch === " ") {
            tokens.push(0);
            continue;
        }
        const v = CHAR_VISEME[ch];
        if (v !== undefined) {
            tokens.push(v);
        }
    }
    if (!tokens.some((v) => v !== 0)) {
        return [];
    }
    const step = durationMs / tokens.length;
    const marks = [];
    let last = -1;
    for (let i = 0; i < tokens.length; i++) {
        const v = tokens[i];
        if (v === last) {
            continue;
        }
        marks.push({ tMs: Math.round(i * step), visemeId: v });
        last = v;
    }
    return marks;
}
function visemeForChar(ch) {
    if (ch === " " || ch === "\n" || ch === "\t" || ch === "\r") {
        return 0;
    }
    return CHAR_VISEME[ch.toLowerCase()];
}
export function visemesFromAlignment(characters, startTimesSeconds) {
    const n = Math.min(characters.length, startTimesSeconds.length);
    const marks = [];
    let last = -1;
    let sawVoiced = false;
    for (let i = 0; i < n; i++) {
        const v = visemeForChar(characters[i] ?? "");
        if (v === undefined) {
            continue;
        }
        if (v !== 0) {
            sawVoiced = true;
        }
        if (v === last) {
            continue;
        }
        marks.push({ tMs: Math.max(0, Math.round((startTimesSeconds[i] ?? 0) * 1000)), visemeId: v });
        last = v;
    }
    return sawVoiced ? marks : [];
}
