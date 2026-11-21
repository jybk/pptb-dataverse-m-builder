import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { generateParameterQuery } from "../src/lib/mcode";
import { buildPowerBiTemplate } from "../src/lib/pbit";
import type { PowerBiProjectInput } from "../src/lib/pbip";
import { URL, setup } from "./modelSetup";

const XML = `<fetch><entity name="account">
    <attribute name="accountid" /><attribute name="name" /><attribute name="industrycode" /><attribute name="new_tags" />
    <attribute name="primarycontactid" /><attribute name="ownerid" /><attribute name="createdby" /><attribute name="modifiedby" />
    <attribute name="revenue" /><attribute name="numberofemployees" /><attribute name="donotemail" /><attribute name="createdon" />
</entity></fetch>`;

function input(opts: Parameters<typeof setup>[1] = {}): { input: PowerBiProjectInput; expected: Record<string, string> } {
    const { data, settings, mainCode } = setup(XML, { lookups: true, overrides: { "|industrycode": { mode: "both" } }, ...opts });
    return {
        input: {
            name: data.mainName,
            main: { name: data.mainName, code: mainCode, columns: data.mainColumns },
            queries: data.queries,
            relationships: data.relationships,
            parameter: settings.useParameter ? { name: settings.parameterName, code: generateParameterQuery(URL, settings) } : null,
        },
        expected: Object.fromEntries([[data.mainName, mainCode], ...data.queries.map((q) => [q.queryName, q.code])]),
    };
}

const utf16 = (bytes: Uint8Array) => new TextDecoder("utf-16le").decode(bytes);

/** PBIP_OUT=dir also writes the templates there, for a check with Microsoft's libraries. */
function writeOut(name: string, bytes: Uint8Array, expected: Record<string, string>) {
    const out = process.env.PBIP_OUT;
    if (!out) return;
    mkdirSync(join(out, `${name}.expected`), { recursive: true });
    writeFileSync(join(out, `${name}.pbit`), bytes);
    for (const [table, code] of Object.entries(expected)) writeFileSync(join(out, `${name}.expected`, `${table}.m`), code.replace(/^\/\/ Suggested query name: .*\n/, ""));
}

describe("Power BI template", () => {
    it("has the parts Desktop writes, in its encodings", async () => {
        const { input: i, expected } = input();
        const bytes = await buildPowerBiTemplate(i);
        writeOut("Template", bytes, expected);
        const zip = await JSZip.loadAsync(bytes);
        expect(Object.keys(zip.files)).toEqual([
            "Version",
            "[Content_Types].xml",
            "DataModelSchema",
            "DiagramLayout",
            "Report/Layout",
            "Settings",
            "Metadata",
            "Report/StaticResources/SharedResources/BaseThemes/CY24SU10.json",
        ]);
        const part = (name: string) => zip.file(name)!.async("uint8array");
        expect(utf16(await part("Version"))).toBe("1.30");
        const types = await part("[Content_Types].xml");
        expect([...types.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
        expect(new TextDecoder().decode(types.slice(3))).toContain('<Override PartName="/DataModelSchema" ContentType="" />');
        for (const name of ["DataModelSchema", "DiagramLayout", "Report/Layout", "Settings", "Metadata"]) {
            const raw = await part(name);
            expect(raw[0]).toBe("{".charCodeAt(0));
            expect(raw[1]).toBe(0);
            expect(() => JSON.parse(utf16(raw))).not.toThrow();
        }
        const layout = JSON.parse(utf16(await part("Report/Layout")));
        expect(layout.sections).toHaveLength(1);
        expect(JSON.parse(layout.config).themeCollection.baseTheme.name).toBe("CY24SU10");
    });

    it("holds the same model as the project, as JSON", async () => {
        const { input: i } = input();
        const zip = await JSZip.loadAsync(await buildPowerBiTemplate(i));
        const db = JSON.parse(utf16(await zip.file("DataModelSchema")!.async("uint8array")));
        expect(db.compatibilityLevel).toBe(1550);
        const model = db.model;
        expect(model.tables.map((t: { name: string }) => t.name)).toEqual([
            "Account",
            "Account Do not allow Emails",
            "Account Industry",
            "Account Tags",
            "Account Tags bridge",
            "Contact",
            "Owner",
            "User",
        ]);
        expect(model.expressions[0]).toMatchObject({ name: "DataverseUrl", kind: "m" });
        expect(model.expressions[0].expression).toContain("IsParameterQuery = true");
        const bridge = model.tables.find((t: { name: string }) => t.name === "Account Tags bridge");
        expect(bridge.isHidden).toBe(true);
        const owner = model.tables.find((t: { name: string }) => t.name === "Owner");
        expect(owner.columns[0]).toMatchObject({ name: "Owner key", isHidden: true, dataType: "string" });
        const account = model.tables[0];
        expect(account.partitions[0].source.expression[0]).toBe("// Generated by Dataverse M Builder for Power Platform ToolBox");
        expect(account.columns.find((c: { name: string }) => c.name === "Annual Revenue")).toMatchObject({ dataType: "decimal", summarizeBy: "sum" });
        expect(model.relationships.find((r: { fromColumn: string }) => r.fromColumn === "Modified By ID")).toMatchObject({ toTable: "User", isActive: false });
        expect(model.relationships.find((r: { fromTable: string; toTable: string }) => r.fromTable === "Account Tags bridge" && r.toTable === "Account")).toMatchObject({
            crossFilteringBehavior: "bothDirections",
        });
    });

    it("is the same file every time for the same model", async () => {
        const a = await buildPowerBiTemplate(input().input);
        const b = await buildPowerBiTemplate(input().input);
        expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true);
    });

    it("has no parameter when the URL is written into the queries", async () => {
        const { input: i, expected } = input({ settings: { useParameter: false } });
        const bytes = await buildPowerBiTemplate(i);
        writeOut("Template written URL", bytes, expected);
        const zip = await JSZip.loadAsync(bytes);
        expect(JSON.parse(utf16(await zip.file("DataModelSchema")!.async("uint8array"))).model.expressions).toBeUndefined();
    });

    it("keeps awkward names as they are", async () => {
        const overrides = { "|industrycode": { mode: "both" as const }, "|name": { name: '"Quoted" name' }, "|revenue": { name: "Revenue: 'net'" } };
        const { input: i, expected } = input({ mainName: "Klant beste", overrides });
        const bytes = await buildPowerBiTemplate(i);
        writeOut("Template names", bytes, expected);
        const zip = await JSZip.loadAsync(bytes);
        const db = JSON.parse(utf16(await zip.file("DataModelSchema")!.async("uint8array")));
        expect(db.model.tables[0].columns.map((c: { name: string }) => c.name)).toContain('"Quoted" name');
    });

    it("holds a lookup table's added columns", async () => {
        const lookupColumns = { contact: { columns: ["emailaddress1"], overrides: {} } };
        const { input: i, expected } = input({ lookupColumns });
        const bytes = await buildPowerBiTemplate(i);
        writeOut("Template extended", bytes, expected);
        const zip = await JSZip.loadAsync(bytes);
        const db = JSON.parse(utf16(await zip.file("DataModelSchema")!.async("uint8array")));
        const contact = db.model.tables.find((t: { name: string }) => t.name === "Contact");
        expect(contact.columns.map((c: { name: string; isHidden?: boolean }) => [c.name, !!c.isHidden])).toEqual([
            ["Contact key", true],
            ["Contact name", false],
            ["Email", false],
        ]);
    });
});
