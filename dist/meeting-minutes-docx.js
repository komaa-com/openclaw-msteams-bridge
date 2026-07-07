/**
 * Deterministic OOXML (.docx) meeting-minutes builder.
 *
 * The end-of-call recap previously asked the model to hand-write an HTML `.doc`, which was neither a
 * real Word document nor deterministic. This builds a minimal but valid OOXML `.docx` entirely in
 * code from (a) the agent's structured summary sections and (b) the speaker-prefixed transcript, so
 * per-speaker attribution is exact and reproducible. We assemble the four fixed parts by hand and zip
 * them with `jszip` (already a dependency — see `src/logging/diagnostic-support-bundle.ts`); the
 * heavier `docx`/`officegen` packages are intentionally avoided.
 */
const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
/** XML-escape text for safe inclusion in an OOXML run. */
function escapeXml(value) {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}
/** A normal-weight paragraph. `xml:space="preserve"` keeps leading/trailing spaces intact. */
function paragraph(text) {
    return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}
/** A bold heading paragraph (no style part needed — bold run + larger size is enough for Word). */
function heading(text) {
    return (`<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
}
/** A document title paragraph (largest). */
function title(text) {
    return (`<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
}
/** A bullet-style paragraph (rendered with a leading "• " — no numbering part needed). */
function bullet(text) {
    return paragraph(`• ${text}`);
}
/** Build the `word/document.xml` body from the structured input. */
function buildDocumentXml(input) {
    const assistantLabel = input.assistantLabel ?? "Assistant";
    const callerLabel = input.callerLabel ?? "Caller";
    const parts = [];
    parts.push(title(input.title));
    if (input.subtitle?.trim()) {
        parts.push(paragraph(input.subtitle.trim()));
    }
    for (const section of input.sections ?? []) {
        const items = section.items.map((i) => i.trim()).filter(Boolean);
        if (items.length === 0) {
            continue;
        }
        parts.push(heading(section.heading));
        for (const item of items) {
            parts.push(bullet(item));
        }
    }
    // Attributed transcript: rendered deterministically from the speaker-prefixed turns. Caller turns
    // already carry a "<Name>: " prefix from the unmixed-audio attribution, so we keep them verbatim;
    // un-prefixed caller turns fall back to the generic caller label, and assistant turns are labelled.
    parts.push(heading("Attributed transcript"));
    for (const turn of input.transcript) {
        const text = turn.text.trim();
        if (!text) {
            continue;
        }
        if (turn.role === "assistant") {
            parts.push(paragraph(`${assistantLabel}: ${text}`));
        }
        else if (/^[^\s:][^:]*:\s/.test(text)) {
            // Already speaker-prefixed (e.g. "Sara: …") — keep the exact attribution.
            parts.push(paragraph(text));
        }
        else {
            parts.push(paragraph(`${callerLabel}: ${text}`));
        }
    }
    return (`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body>${parts.join("")}` +
        `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
        `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>` +
        `</w:sectPr></w:body></w:document>`);
}
const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `</Types>`;
const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;
const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
/** MIME type for a `.docx` (also mapped in `packages/media-core/src/mime.ts`). */
export const MINUTES_DOCX_MIME = DOCX_CONTENT_TYPE;
/**
 * Build a minimal valid OOXML `.docx` as a Buffer. The bytes are produced entirely in code, so the
 * document — including per-speaker attribution — is deterministic for a given input.
 */
export async function buildMinutesDocx(input) {
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES_XML);
    zip.file("_rels/.rels", ROOT_RELS_XML);
    zip.file("word/document.xml", buildDocumentXml(input));
    zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS_XML);
    return zip.generateAsync({ type: "nodebuffer" });
}
