import * as PQP from "@microsoft/powerquery-parser";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defaultModeFor, modeOptionsFor, resolveOutputColumns } from "../src/lib/columns";
import { loadAttributes } from "../src/lib/dataverse";
import { buildQueryModel, parseFetchXml } from "../src/lib/fetchxml";
import { generateFetchXmlQuery, mIdent } from "../src/lib/mcode";
import { mapAttributes } from "../src/lib/metadata";
import { buildPreview } from "../src/lib/preview";
import { DebouncedSaver } from "../src/lib/saver";
import { DEFAULT_SETTINGS, type BuilderSettings } from "../src/lib/types";
import { fixture, rawAccountAttributes, rawAccountDates } from "./fixtures";

const URL = "https://contoso.crm4.dynamics.com";

function build(xml: string, opts: { excluded?: string[]; includeRecordUrl?: boolean; withoutMetadataFor?: string } = {}) {
    const { tables, attributes } = fixture();
    if (opts.withoutMetadataFor) attributes.delete(opts.withoutMetadataFor);
    return buildQueryModel({
        doc: parseFetchXml(xml),
        tables,
        attributes,
        pageSize: 5000,
        includeRecordUrl: !!opts.includeRecordUrl,
        excluded: new Set(opts.excluded ?? []),
    });
}

async function expectValidM(code: string) {
    const result = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, code);
    if (!PQP.TaskUtils.isParseStageOk(result)) throw new Error(`Generated M doesn't parse:\n\n${code}`);
}

describe("table-name columns (EntityName)", () => {
    const [regarding] = mapAttributes([
        {
            LogicalName: "objecttypecode",
            SchemaName: "ObjectTypeCode",
            DisplayName: { UserLocalizedLabel: { Label: "Object Type" } },
            AttributeType: "EntityName",
            AttributeTypeName: { Value: "EntityNameType" },
            IsValidForRead: true,
        },
    ]);

    it("come out as the table's logical name, without a label Dataverse doesn't send", () => {
        expect(regarding.kind).toBe("entityname");
        const spec = {
            id: "|objecttypecode",
            key: "objecttypecode",
            entity: "annotation",
            linkAlias: null,
            linkLabel: null,
            attrAlias: null,
            aggregate: null,
            dategrouping: null,
            attribute: regarding,
        };
        expect(modeOptionsFor(spec)).toBeNull();
        expect(defaultModeFor(spec, DEFAULT_SETTINGS)).toBeNull();
        const outputs = resolveOutputColumns([spec], DEFAULT_SETTINGS, { "|objecttypecode": { mode: "label" } });
        expect(outputs).toHaveLength(1);
        expect(outputs[0]).toMatchObject({ sourceKey: "objecttypecode", mType: "type text" });
    });
});

describe("missing metadata", () => {
    afterEach(() => {
        delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
    });

    function mockApi(failDates: boolean) {
        (window as unknown as { dataverseAPI: unknown }).dataverseAPI = {
            getEntityRelatedMetadata: async () => ({ value: rawAccountAttributes }),
            queryData: async (q: string) => {
                if (q.includes("DateTimeAttributeMetadata")) {
                    if (failDates) throw new Error("403 Forbidden");
                    return { value: rawAccountDates };
                }
                return { value: [] };
            },
        };
    }

    it("reports date formats it couldn't read, since date-only columns then look like date and time", async () => {
        mockApi(true);
        const failed = await loadAttributes("account");
        expect(failed.dateFormatsMissing).toBe(true);
        expect(failed.attributes.find((a) => a.logicalName === "new_startdate")?.dateBehavior).toBeUndefined();

        mockApi(false);
        const ok = await loadAttributes("account");
        expect(ok.dateFormatsMissing).toBe(false);
        expect(ok.attributes.find((a) => a.logicalName === "new_startdate")?.dateBehavior).toBe("DateOnly");
    });

    it("says <all-attributes /> is left out when the table's columns couldn't be loaded", () => {
        const xml = `<fetch><entity name="account"><attribute name="name" />
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><all-attributes /></link-entity></entity></fetch>`;
        const model = build(xml, { withoutMetadataFor: "contact" });
        expect(model.warnings.join(" ")).toContain("<all-attributes /> on c is left out because that table's columns couldn't be loaded");
        expect(model.warnings.join(" ")).not.toContain("replaced with its 0");
        expect(model.fetchXml).not.toContain("all-attributes");
    });
});

describe("paging", () => {
    it("treats a paging cookie without a pagingcookie attribute as no cookie", async () => {
        const model = build(`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`);
        const code = generateFetchXmlQuery(model, resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {}), DEFAULT_SETTINGS, URL);
        await expectValidM(code);
        expect(code).toContain('RawCookie = if CookieXml = null then "" else Text.BetweenDelimiters(CookieXml, "pagingcookie=""", """"),');
        expect(code).toContain('NextCookie = if RawCookie = "" then null else UrlDecode(UrlDecode(RawCookie))');
    });

    it("uses a whole number from count as the page size", () => {
        const model = build(`<fetch count="2.5"><entity name="account"><attribute name="name" /></entity></fetch>`);
        expect(model.pageSize).toBe(2);
        expect(build(`<fetch count="0.5"><entity name="account"><attribute name="name" /></entity></fetch>`).pageSize).toBe(5000);
    });

    it("keeps the FetchXML starting with <fetch, which the paging step builds on", () => {
        const model = build(
            `<?xml version="1.0"?>\n<!-- a view --><fetch distinct="true"><entity name="account"><attribute name="name" /><order attribute="name" /></entity></fetch>`,
        );
        expect(model.fetchLines[0].text.startsWith("<fetch")).toBe(true);
        expect(model.fetchXml.startsWith("<fetch")).toBe(true);
    });
});

describe("unticked columns in a distinct query", () => {
    it("keeps a linked column that the main table sorts on by the link's alias", () => {
        const xml = `<fetch distinct="true"><entity name="account"><attribute name="name" />
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="fullname" /><attribute name="emailaddress1" /></link-entity>
            <order entityname="c" attribute="fullname" /></entity></fetch>`;
        const model = build(xml, { excluded: ["c|fullname", "c|emailaddress1"] });
        expect(model.fetchXml).toContain('<attribute name="fullname" />');
        expect(model.fetchXml).not.toContain("emailaddress1");
    });

    it("still removes a main-table column with the same name as the sorted linked one", () => {
        const xml = `<fetch distinct="true"><entity name="account"><attribute name="name" /><attribute name="telephone1" />
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="telephone1" /></link-entity>
            <order entityname="c" attribute="telephone1" /></entity></fetch>`;
        const model = build(xml, { excluded: ["|telephone1"] });
        expect(model.fetchXml.match(/<attribute name="telephone1" \/>/g)).toHaveLength(1);
    });
});

describe("record URLs", () => {
    const settings: BuilderSettings = { ...DEFAULT_SETTINGS, includeRecordUrl: true, appId: "a b&c" };

    it("encode the app id", async () => {
        const model = build(`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`, { includeRecordUrl: true });
        const outputs = resolveOutputColumns(model.columns, settings, {}, !!model.recordUrlKey);
        const code = generateFetchXmlQuery(model, outputs, settings, URL);
        await expectValidM(code);
        expect(code).toContain("appid=a%20b%26c&");
        const preview = buildPreview(model, outputs, [{ name: "Contoso", accountid: "00000000-0000-0000-0000-000000000001" }], settings, URL);
        expect(String(preview.rows[0][preview.columns.indexOf("Record URL")])).toContain("appid=a%20b%26c&");
    });

    it("only reserve the Record URL name when the query has one", () => {
        // An aggregate query has no record URL, so a column named "Record URL" keeps its name
        const model = build(`<fetch aggregate="true"><entity name="account"><attribute name="name" alias="Record URL" groupby="true" /></entity></fetch>`, {
            includeRecordUrl: true,
        });
        expect(model.recordUrlKey).toBeNull();
        expect(resolveOutputColumns(model.columns, settings, {}, false).map((o) => o.name)).toEqual(["Record URL"]);
        expect(resolveOutputColumns(model.columns, settings, {}, true).map((o) => o.name)).toEqual(["Record URL (2)"]);
    });
});

describe("M keywords as names", () => {
    it("quotes column names that are M keywords", async () => {
        expect(mIdent("type")).toBe('#"type"');
        expect(mIdent("let")).toBe('#"let"');
        expect(mIdent("Name")).toBe("Name");
        const model = build(`<fetch><entity name="account"><attribute name="name" alias="type" /><attribute name="telephone1" alias="let" /></entity></fetch>`);
        const code = generateFetchXmlQuery(model, resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {}), DEFAULT_SETTINGS, URL);
        await expectValidM(code);
    });
});

describe("saving with a delay", () => {
    beforeEach(() => vi.useFakeTimers());
    afterEach(() => vi.useRealTimers());

    it("writes once, after the last change", async () => {
        const write = vi.fn(async () => undefined);
        const saver = new DebouncedSaver(write, 800);
        saver.schedule("session:a", 1);
        saver.schedule("session:a", 2);
        vi.advanceTimersByTime(799);
        expect(write).not.toHaveBeenCalled();
        vi.advanceTimersByTime(1);
        expect(write).toHaveBeenCalledTimes(1);
        expect(write).toHaveBeenCalledWith("session:a", 2);
    });

    it("doesn't drop a waiting save when another environment's value comes in", () => {
        const write = vi.fn(async () => undefined);
        const saver = new DebouncedSaver(write, 800);
        saver.schedule("session:a", "last edit in A");
        saver.schedule("session:b", "B");
        expect(write).toHaveBeenCalledWith("session:a", "last edit in A");
        vi.advanceTimersByTime(800);
        expect(write).toHaveBeenLastCalledWith("session:b", "B");
        expect(write).toHaveBeenCalledTimes(2);
    });

    it("writes what's waiting when the tool closes, and nothing twice", () => {
        const write = vi.fn(async () => undefined);
        const saver = new DebouncedSaver(write, 800);
        saver.flush();
        expect(write).not.toHaveBeenCalled();
        saver.schedule("builderSettings", { naming: "logical" });
        saver.flush();
        vi.advanceTimersByTime(2000);
        saver.flush();
        expect(write).toHaveBeenCalledTimes(1);
    });

    it("keeps working when a write fails", async () => {
        const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
        const saver = new DebouncedSaver(async () => Promise.reject(new Error("disk full")), 10);
        saver.schedule("k", 1);
        saver.flush();
        await vi.runAllTimersAsync();
        expect(warn).toHaveBeenCalled();
        warn.mockRestore();
    });
});
