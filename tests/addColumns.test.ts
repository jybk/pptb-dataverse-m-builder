import { describe, expect, it } from "vitest";
import { resolveOutputColumns } from "../src/lib/columns";
import { buildQueryModel, parseFetchXml, type ModelInput } from "../src/lib/fetchxml";
import { generateFetchXmlQuery } from "../src/lib/mcode";
import { DEFAULT_SETTINGS } from "../src/lib/types";
import { fixture } from "./fixtures";
import { URL, expectValidM } from "./modelSetup";

const VIEW = `<fetch>
  <entity name="account">
    <attribute name="name" />
    <link-entity name="contact" from="contactid" to="primarycontactid" alias="pc" link-type="outer">
      <attribute name="fullname" />
    </link-entity>
    <link-entity name="contact" from="parentcustomerid" to="accountid" link-type="outer" />
    <filter><condition attribute="statecode" operator="eq" value="0" /></filter>
  </entity>
</fetch>`;

function model(xml: string, added: string[] = [], excluded: string[] = []) {
    const { tables, attributes } = fixture();
    const input: ModelInput = { doc: parseFetchXml(xml), tables, attributes, pageSize: 5000, includeRecordUrl: false, added: new Set(added), excluded: new Set(excluded) };
    return buildQueryModel(input);
}

describe("columns added to a FetchXML query", () => {
    it("lists the main table and the linked tables that return columns", () => {
        const m = model(VIEW);
        expect(m.sources).toEqual([
            { alias: null, entity: "account", label: "Account" },
            { alias: "pc", entity: "contact", label: "pc" },
        ]);
    });

    it("go into the main table or the linked table, and into the generated query", async () => {
        const m = model(VIEW, ["|revenue", "pc|emailaddress1"]);
        const added = m.columns.filter((c) => c.added).map((c) => c.id);
        expect(added).toEqual(["|revenue", "pc|emailaddress1"]);
        expect(m.columns.find((c) => c.id === "pc|emailaddress1")).toMatchObject({ entity: "contact", linkAlias: "pc", key: "pc.emailaddress1" });
        // Attributes before the link, as in any FetchXML
        expect(m.fetchXml).toMatch(/<attribute name="name" \/>\s*<attribute name="revenue" \/>\s*<link-entity/);
        expect(m.fetchXml).toMatch(/<attribute name="fullname" \/>\s*<attribute name="emailaddress1" \/>\s*<\/link-entity>/);
        const outputs = resolveOutputColumns(m.columns, DEFAULT_SETTINGS, {}, false);
        expect(outputs.map((o) => o.name)).toEqual(["Account Name", "Annual Revenue", "pc Full Name", "pc Email"]);
        await expectValidM(generateFetchXmlQuery(m, outputs, DEFAULT_SETTINGS, URL));
    });

    it("skip columns that are already there, that the table doesn't have, or of tables the query doesn't return", () => {
        const m = model(VIEW, ["|name", "|nonexistent", "nope|fullname", "contact|emailaddress1"]);
        expect(m.columns.filter((c) => c.added)).toEqual([]);
        expect(m.columns.map((c) => c.id)).toEqual(["|name", "pc|fullname"]);
    });

    it("can be unticked again like any other column", () => {
        const m = model(VIEW, ["|revenue"], ["|revenue"]);
        expect(m.fetchXml).not.toContain("revenue");
    });

    it("aren't offered or added for aggregate queries, where they'd change the grouping", () => {
        const m = model(`<fetch aggregate="true"><entity name="account"><attribute name="industrycode" groupby="true" alias="industry" /><attribute name="accountid" aggregate="count" alias="n" /></entity></fetch>`, [
            "|revenue",
        ]);
        expect(m.sources).toEqual([]);
        expect(m.fetchXml).not.toContain("revenue");
    });
});
