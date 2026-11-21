import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { generateFetchXmlQuery } from "../src/lib/mcode";
import { buildWorkbook } from "../src/lib/xlsx";
import { URL, setup } from "./modelSetup";

const XML = `<fetch><entity name="account"><attribute name="accountid" /><attribute name="name" /><attribute name="industrycode" /><attribute name="revenue" /></entity></fetch>`;

/** The Power Query package inside a workbook: base64 in customXml, a length-prefixed zip inside (MS-QDEFF). */
async function mashup(zip: JSZip): Promise<JSZip> {
    // Excel writes this part in UTF-16, as the library does
    let base64 = "";
    for (const name of Object.keys(zip.files).filter((f) => /^customXml\/item\d+\.xml$/.test(f))) {
        const raw = await zip.file(name)!.async("uint8array");
        const xml = new TextDecoder(raw[0] === 0xff && raw[1] === 0xfe ? "utf-16le" : "utf-8").decode(raw);
        base64 = /<DataMashup[^>]*>([^<]+)<\/DataMashup>/.exec(xml)?.[1] ?? base64;
    }
    const bytes = Buffer.from(base64, "base64");
    const length = bytes.readUInt32LE(4);
    return JSZip.loadAsync(bytes.subarray(8, 8 + length));
}

describe("Excel workbook", () => {
    it("holds the main query, loaded to a table and refreshed on open", async () => {
        const { model, outputs, settings } = setup(XML);
        const code = generateFetchXmlQuery(model, outputs, { ...settings, useParameter: false }, URL);
        const bytes = await buildWorkbook("Accounts", code, outputs.map((o) => o.name));
        const zip = await JSZip.loadAsync(bytes);
        const section = await (await mashup(zip)).file("Formulas/Section1.m")!.async("string");
        expect(section).toContain('shared #"Accounts" =');
        expect(section).toContain('DataverseUrl = "https://contoso.crm4.dynamics.com/api/data/v9.2",');
        expect(section).not.toContain("Suggested query name");
        const table = await zip.file(Object.keys(zip.files).find((f) => /^xl\/tables\/table\d+\.xml$/.test(f))!)!.async("string");
        for (const o of outputs) expect(table).toContain(`name="${o.name}"`);
        const connections = await zip.file("xl/connections.xml")!.async("string");
        expect(connections).toContain("Location=&quot;Accounts&quot;;");
        expect(connections).toContain('command="SELECT * FROM [Accounts]" refreshOnLoad="1"');
        expect(await zip.file("xl/queryTables/queryTable1.xml")!.async("string")).toContain('refreshOnLoad="1"');
        const core = await zip.file("docProps/core.xml")!.async("string");
        expect(core).toContain("<dc:title>Accounts</dc:title>");
    });
});
