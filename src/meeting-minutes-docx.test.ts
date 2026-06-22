import { describe, expect, it } from "vitest";
import { buildMinutesDocx } from "./meeting-minutes-docx.js";

/** Unzip a generated .docx in-memory and return the bytes of an inner part. */
async function readZipEntry(buffer: Buffer, entry: string): Promise<string> {
  const { default: JSZip } = await import("jszip");
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(entry);
  if (!file) {
    throw new Error(`missing ${entry}`);
  }
  return file.async("string");
}

describe("buildMinutesDocx", () => {
  it("returns a ZIP-magic buffer containing word/document.xml", async () => {
    const buffer = await buildMinutesDocx({
      title: "Meeting minutes",
      subtitle: "Call with Sara — ~5 min, 2 human participant(s).",
      transcript: [{ role: "user", text: "Sara: let's ship friday" }],
    });
    expect(Buffer.isBuffer(buffer)).toBe(true);
    // ZIP local file header magic "PK\x03\x04".
    expect(buffer.subarray(0, 4)).toEqual(Buffer.from([0x50, 0x4b, 0x03, 0x04]));

    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(buffer);
    expect(zip.file("word/document.xml")).toBeTruthy();
    expect(zip.file("[Content_Types].xml")).toBeTruthy();
    expect(zip.file("_rels/.rels")).toBeTruthy();
    expect(zip.file("word/_rels/document.xml.rels")).toBeTruthy();
  });

  it("renders a speaker-prefixed turn attributed in the document XML", async () => {
    const buffer = await buildMinutesDocx({
      title: "Meeting minutes",
      transcript: [
        { role: "user", text: "Sara: we should raise the budget" },
        { role: "assistant", text: "Noted." },
        { role: "user", text: "no name here" },
      ],
    });
    const xml = await readZipEntry(buffer, "word/document.xml");
    // The exact speaker attribution is preserved verbatim from the prefixed turn.
    expect(xml).toContain("Sara: we should raise the budget");
    // Assistant turns are labelled; un-prefixed caller turns fall back to the generic label.
    expect(xml).toContain("Assistant: Noted.");
    expect(xml).toContain("Caller: no name here");
  });

  it("renders headed sections with bullets and skips empty ones", async () => {
    const buffer = await buildMinutesDocx({
      title: "Meeting minutes",
      sections: [
        { heading: "Key points", items: ["budget is on track"] },
        { heading: "Decisions", items: [] },
      ],
      transcript: [],
    });
    const xml = await readZipEntry(buffer, "word/document.xml");
    expect(xml).toContain("Key points");
    expect(xml).toContain("budget is on track");
    // An empty section produces no heading.
    expect(xml).not.toContain("Decisions");
  });

  it("XML-escapes special characters", async () => {
    const buffer = await buildMinutesDocx({
      title: "Minutes <&>",
      transcript: [{ role: "user", text: 'Sara: a & b < c > d "e"' }],
    });
    const xml = await readZipEntry(buffer, "word/document.xml");
    expect(xml).toContain("Minutes &lt;&amp;&gt;");
    expect(xml).toContain("a &amp; b &lt; c &gt; d &quot;e&quot;");
    expect(xml).not.toContain("a & b");
  });
});
