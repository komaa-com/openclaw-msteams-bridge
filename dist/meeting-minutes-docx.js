const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
function escapeXml(value) {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}
function paragraph(text) {
    return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}
function heading(text) {
    return (`<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
}
function title(text) {
    return (`<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
}
function bullet(text) {
    return paragraph(`• ${text}`);
}
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
export const MINUTES_DOCX_MIME = DOCX_CONTENT_TYPE;
export async function buildMinutesDocx(input) {
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES_XML);
    zip.file("_rels/.rels", ROOT_RELS_XML);
    zip.file("word/document.xml", buildDocumentXml(input));
    zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS_XML);
    return zip.generateAsync({ type: "nodebuffer" });
}
