import * as PQP from "@microsoft/powerquery-parser";
import { describe, expect, it } from "vitest";
import { resolveOutputColumns } from "../src/lib/columns";
import { buildQueryModel, buildTableFetchXml, fetchEntities, parseFetchXml, previewFetchXml } from "../src/lib/fetchxml";
import { mapOptions } from "../src/lib/metadata";
import { choiceColumns, generateChoiceQuery, generateFetchXmlQuery, generateODataQuery, generateParameterQuery, mText, normaliseUrl, ALLOWED_AUTH_HEADERS } from "../src/lib/mcode";
import { buildPreview } from "../src/lib/preview";
import { DEFAULT_SETTINGS, type BuilderSettings, type ColumnOverrides } from "../src/lib/types";
import { fixture } from "./fixtures";

const URL = "https://contoso.crm4.dynamics.com/";

async function expectValidM(code: string) {
    const result = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, code);
    if (!PQP.TaskUtils.isParseStageOk(result)) {
        const err = (result as { error?: { message?: string } }).error?.message;
        throw new Error(`Generated M doesn't parse: ${err}\n\n${code}`);
    }
}

function tableModel(attrs: string[], settings: BuilderSettings = DEFAULT_SETTINGS) {
    const { tables, attributes } = fixture();
    const table = tables.get("account")!;
    const xml = buildTableFetchXml(table, attrs, settings.activeOnly);
    return buildQueryModel({
        doc: parseFetchXml(xml),
        tables,
        attributes,
        pageSize: settings.pageSize,
        includeRecordUrl: settings.includeRecordUrl,
        fromTablePicker: true,
        odataFilter: settings.activeOnly ? "statecode eq 0" : null,
    });
}

describe("metadata mapping", () => {
    it("keeps queryable tables only and sorts by display name", () => {
        const { tableList } = fixture();
        expect(tableList.map((t) => t.logicalName)).toEqual(["account", "contact", "new_project"]);
    });

    it("drops helper, image and unreadable columns and reads date behaviour", () => {
        const { attributes } = fixture();
        const names = attributes.get("account")!.map((a) => a.logicalName);
        expect(names[0]).toBe("accountid");
        expect(names[1]).toBe("name");
        expect(names).not.toContain("primarycontactidname");
        expect(names).not.toContain("entityimage");
        expect(names).not.toContain("hidden");
        const start = attributes.get("account")!.find((a) => a.logicalName === "new_startdate")!;
        expect(start.dateBehavior).toBe("DateOnly");
        expect(attributes.get("account")!.find((a) => a.logicalName === "ownerid")!.targets).toEqual(["systemuser", "team"]);
    });

    it("maps option metadata", () => {
        expect(mapOptions("choice", { OptionSet: null, GlobalOptionSet: { Options: [{ Value: 1, Label: { UserLocalizedLabel: { Label: "Accounting" } } }] } })).toEqual([
            { value: 1, label: "Accounting" },
        ]);
        expect(
            mapOptions("boolean", {
                OptionSet: {
                    TrueOption: { Value: 1, Label: { UserLocalizedLabel: { Label: "Do Not Allow" } } },
                    FalseOption: { Value: 0, Label: { LocalizedLabels: [{ Label: "Allow" }] } },
                },
            }),
        ).toEqual([
            { value: true, label: "Do Not Allow" },
            { value: false, label: "Allow" },
        ]);
    });
});

describe("table picker queries", () => {
    const attrs = [
        "accountid",
        "name",
        "primarycontactid",
        "industrycode",
        "revenue",
        "createdon",
        "new_startdate",
        "new_reviewdate",
        "donotemail",
        "numberofemployees",
        "new_tags",
    ];

    it("predicts the Web API keys", () => {
        const model = tableModel(attrs);
        const keys = Object.fromEntries(model.columns.map((c) => [c.attribute.logicalName, c.key]));
        expect(keys.primarycontactid).toBe("_primarycontactid_value");
        expect(keys.name).toBe("name");
        expect(model.fetchXml).toContain('<condition attribute="statecode" operator="eq" value="0" />');
        expect(model.singleRequest).toBe(false);
        expect(model.odata?.select).toContain("_primarycontactid_value");
    });

    it("resolves output columns with default modes", () => {
        const model = tableModel(attrs);
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        const byName = Object.fromEntries(out.map((o) => [o.name, o]));
        expect(byName["Account ID"].sourceKey).toBe("accountid");
        expect(byName["Primary Contact"].sourceKey).toBe("_primarycontactid_value@OData.Community.Display.V1.FormattedValue");
        expect(byName["Primary Contact ID"].sourceKey).toBe("_primarycontactid_value");
        expect(byName["Industry"].sourceKey).toBe("industrycode@OData.Community.Display.V1.FormattedValue");
        expect(byName["Annual Revenue"].mType).toBe("Currency.Type");
        expect(byName["Created On"].conversion).toBe("utc");
        expect(byName["Start Date"].conversion).toBe("date");
        expect(byName["Start Date"].mType).toBe("type date");
        expect(byName["Review Date"].conversion).toBe("wallclock");
        expect(byName["Do not allow Emails"].mType).toBe("type logical");
        expect(byName["Tags"].sourceKey).toContain("FormattedValue");
    });

    it("generates valid, paged FetchXML-based M", async () => {
        const model = tableModel(attrs);
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        const code = generateFetchXmlQuery(model, out, DEFAULT_SETTINGS, URL);
        await expectValidM(code);
        expect(code).toContain('EntitySetName = "accounts"');
        expect(code).toContain("Web.Contents(DataverseUrl, [");
        expect(code).toContain("RelativePath = EntitySetName,");
        expect(code).not.toContain('api/data/v9.2/" &');
        expect(code).toContain("Microsoft.Dynamics.CRM.fetchxmlpagingcookie");
        expect(code).toContain('"<entity name=""account"">"');
        expect(code).toContain("ToDate = ");
        expect(code).toContain("// Needs a text parameter named DataverseUrl = https://contoso.crm4.dynamics.com/api/data/v9.2 (no trailing slash)");
        expect(code).not.toContain("lookuplogicalname");
    });

    it("generates valid OData M", async () => {
        const model = tableModel(attrs);
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        const code = generateODataQuery(model, out, DEFAULT_SETTINGS, URL);
        await expectValidM(code);
        expect(code).toContain('#"$select" = "accountid,name,_primarycontactid_value,');
        expect(code).toContain('#"$filter" = "statecode eq 0"');
        expect(code).toContain("odata.maxpagesize=5000");
        expect(code).toContain("Web.Contents(DataverseUrl, [");
        expect(code).toContain("@odata.nextLink");
    });

    it("supports every option combination", async () => {
        const settings: BuilderSettings = {
            ...DEFAULT_SETTINGS,
            useParameter: false,
            includeRecordUrl: true,
            appId: "11111111-2222-3333-4444-555555555555",
            naming: "logical",
            lookupMode: "all",
            choiceMode: "both",
            booleanMode: "label",
            dateMode: "offset",
            utcOffsetMinutes: -330,
            activeOnly: false,
        };
        const model = tableModel(["name", "primarycontactid", "ownerid", "industrycode", "createdon", "donotemail"], settings);
        expect(model.recordUrlKey).toBe("accountid");
        const out = resolveOutputColumns(model.columns, settings, {});
        const names = out.map((o) => o.name);
        expect(names).toEqual(expect.arrayContaining(["primarycontactidname", "primarycontactid", "primarycontactidtype", "industrycodename", "industrycode", "donotemailname"]));
        for (const gen of [generateFetchXmlQuery, generateODataQuery]) {
            const code = gen(model, out, settings, URL);
            await expectValidM(code);
            expect(code).toContain('DataverseUrl = "https://contoso.crm4.dynamics.com/api/data/v9.2"');
            expect(code).toContain('Text.PositionOf(Text.Lower(DataverseUrl), "/api/data/")');
            expect(code).toContain("lookuplogicalname");
            expect(code).toContain("SwitchZone(DateTimeZone.FromText(value), -5, -30)");
            expect(code).toContain("appid=11111111-2222-3333-4444-555555555555&pagetype=entityrecord&etn=account&id=");
            // Primary key isn't an output column here, so the helper column is dropped at the end
            expect(code).toContain("Result = Table.SelectColumns(");
            expect(code).not.toContain("statecode eq 0");
        }
    });

    it("honours overrides and keeps names unique", async () => {
        const model = tableModel(["name", "primarycontactid", "industrycode"]);
        const overrides: ColumnOverrides = {
            "|primarycontactid": { mode: "name", name: "Account Name" },
            "|industrycode": { mode: "value" },
        };
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, overrides);
        expect(out.map((o) => o.name)).toEqual(["Account Name", "Account Name (2)", "Industry"]);
        expect(out[2].mType).toBe("Int64.Type");
        await expectValidM(generateFetchXmlQuery(model, out, DEFAULT_SETTINGS, URL));
    });
});

describe("custom FetchXML", () => {
    const fetch = `<?xml version="1.0"?>
<fetch version="1.0" mapping="logical" count="250" page="3" paging-cookie="x">
  <!-- linked contact -->
  <entity name="account">
    <attribute name="name" />
    <attribute name="createdon" alias="Created" />
    <link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer">
      <attribute name="fullname" />
      <attribute name="parentcustomerid" />
      <filter><condition attribute="emailaddress1" operator="like" value="%#(test)&amp;&quot;x&quot;%" /></filter>
    </link-entity>
    <link-entity name="contact" from="parentcustomerid" to="accountid" link-type="exists">
      <attribute name="fullname" />
    </link-entity>
    <link-entity name="contact" from="parentcustomerid" to="accountid" alias="kids">
      <all-attributes />
    </link-entity>
  </entity>
</fetch>`;

    it("lists the tables to load metadata for", () => {
        expect(fetchEntities(parseFetchXml(fetch))).toEqual({ root: "account", all: ["account", "contact"] });
    });

    it("assigns aliases, expands all-attributes and computes keys", () => {
        const { tables, attributes } = fixture();
        const model = buildQueryModel({ doc: parseFetchXml(fetch), tables, attributes, pageSize: 5000, includeRecordUrl: false });
        expect(model.fetchXml).toContain('alias="contact"');
        expect(model.fetchXml).not.toContain("page=");
        expect(model.fetchXml).not.toContain("count=");
        expect(model.fetchXml).not.toContain("all-attributes");
        expect(model.pageSize).toBe(250);
        const keys = model.columns.map((c) => c.key);
        expect(keys).toContain("contact.fullname");
        expect(keys).toContain("contact.parentcustomerid");
        expect(keys).toContain("Created");
        expect(keys).toContain("kids.emailaddress1");
        // columns under the exists link aren't returned
        expect(keys.filter((k) => k.endsWith(".fullname"))).toEqual(["contact.fullname", "kids.fullname"]);
        expect(model.odata).toBeNull();
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        // Linked through the Primary Contact lookup, so its columns are named after the lookup
        expect(out.map((o) => o.name)).toContain("Primary Contact Full Name");
        expect(out.map((o) => o.name)).toContain("kids Full Name");
    });

    it("tidies a view's FetchXML: readable aliases, no designer attributes, main columns first", async () => {
        const { tables, attributes } = fixture();
        const view = `<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" no-lock="false" userqueryid="dcdaa275-db7a-f111-ab0e-3833c5ceec57">
  <entity name="account">
    <attribute name="name" />
    <filter type="and">
      <condition entityname="a_48f82b1a8ad844bd90d915e7b3c4f263" attribute="emailaddress1" operator="not-null" />
    </filter>
    <link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="a_48f82b1a8ad844bd90d915e7b3c4f263" visible="false">
      <attribute name="emailaddress1" />
      <attribute name="parentcustomerid" />
    </link-entity>
    <link-entity name="contact" from="parentcustomerid" to="accountid" link-type="outer" alias="a_0123456789abcdef0123456789abcdef">
      <attribute name="fullname" />
    </link-entity>
    <attribute name="createdon" />
    <order attribute="createdon" descending="true" />
  </entity>
</fetch>`;
        const model = buildQueryModel({ doc: parseFetchXml(view), tables, attributes, pageSize: 5000, includeRecordUrl: false });
        expect(model.fetchLines[0].text).toBe('<fetch distinct="true">');
        expect(model.fetchXml).not.toMatch(/a_[0-9a-f]{32}|visible=|userqueryid|output-format|no-lock|mapping=|version=/);
        expect(model.fetchXml).toContain('<condition entityname="contact" attribute="emailaddress1" operator="not-null" />');
        expect(model.fetchXml).toContain('alias="contact2"');
        // attributes moved before the filter, everything else keeps its order
        const lines = model.fetchLines.map((l) => l.text);
        expect(lines.indexOf('<attribute name="createdon" />')).toBeLessThan(lines.findIndex((l) => l.startsWith("<filter")));
        expect(lines.findIndex((l) => l.startsWith("<order"))).toBeGreaterThan(lines.findIndex((l) => l.startsWith("<link-entity")));
        // main table first, then each linked table
        expect(model.columns.map((c) => c.key)).toEqual(["name", "createdon", "contact.emailaddress1", "contact.parentcustomerid", "contact2.fullname"]);
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        expect(out.map((o) => o.name)).toEqual([
            "Account Name",
            "Created On",
            "Primary Contact Email",
            "Primary Contact Company Name",
            "Primary Contact Company Name ID",
            // one-to-many link: named after the linked table
            "Contact Full Name",
        ]);
        const code = generateFetchXmlQuery(model, out, DEFAULT_SETTINGS, URL);
        await expectValidM(code);
        expect(code).toContain('Formatted = "@OData.Community.Display.V1.FormattedValue",');
        expect(code).toMatch(/\{"contact\.parentcustomerid" & Formatted, +"Primary Contact Company Name", +type text\}/);
        expect(code).toMatch(/\{"createdon", +"Created On", +type datetime, ToDateTime\}/);
        expect(code).toContain("Renamed = Table.RenameColumns(Source, List.Transform(Columns, each {_{0}, _{1}})),");
        expect(code).toContain("Typed = Table.TransformColumnTypes(Converted, List.Transform(Columns, each {_{1}, _{2}}))");
    });

    it("produces valid M even with awkward characters in the FetchXML", async () => {
        const { tables, attributes } = fixture();
        const model = buildQueryModel({ doc: parseFetchXml(fetch), tables, attributes, pageSize: 5000, includeRecordUrl: true });
        const out = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {});
        const code = generateFetchXmlQuery(model, out, DEFAULT_SETTINGS, URL);
        await expectValidM(code);
        expect(code).toContain("#(#)(test)");
        expect(code).toContain("&amp;&quot;x&quot;");
    });

    it("sends top and aggregate queries as one request", async () => {
        const { tables, attributes } = fixture();
        const top = buildQueryModel({
            doc: parseFetchXml(`<fetch top="10"><entity name="account"><attribute name="name"/></entity></fetch>`),
            tables,
            attributes,
            pageSize: 5000,
            includeRecordUrl: false,
        });
        expect(top.singleRequest).toBe(true);
        const agg = buildQueryModel({
            doc: parseFetchXml(
                `<fetch aggregate="true"><entity name="account"><attribute name="industrycode" groupby="true" alias="Industry"/><attribute name="accountid" aggregate="count" alias="Accounts"/><attribute name="createdon" groupby="true" dategrouping="year" alias="Year"/></entity></fetch>`,
            ),
            tables,
            attributes,
            pageSize: 5000,
            includeRecordUrl: true,
        });
        expect(agg.singleRequest).toBe(true);
        expect(agg.recordUrlKey).toBeNull();
        const out = resolveOutputColumns(agg.columns, DEFAULT_SETTINGS, {});
        expect(out.find((o) => o.name === "Accounts")?.mType).toBe("Int64.Type");
        expect(out.find((o) => o.name === "Year")?.mType).toBe("Int64.Type");
        expect(out.find((o) => o.name === "Industry")?.role).toBe("label");
        for (const m of [top, agg]) {
            const code = generateFetchXmlQuery(m, resolveOutputColumns(m.columns, DEFAULT_SETTINGS, {}), DEFAULT_SETTINGS, URL);
            await expectValidM(code);
            expect(code).not.toContain("GetPage");
        }
    });

    it("flags unknown columns and bad XML", () => {
        const { tables, attributes } = fixture();
        const model = buildQueryModel({
            doc: parseFetchXml(`<fetch><entity name="account"><attribute name="nope"/></entity></fetch>`),
            tables,
            attributes,
            pageSize: 5000,
            includeRecordUrl: false,
        });
        expect(model.columns[0].attribute.kind).toBe("unknown");
        expect(model.warnings.join(" ")).toContain('"nope"');
        expect(() => parseFetchXml("<fetch><entity>")).toThrow(/valid XML/);
        expect(() => parseFetchXml("<foo/>")).toThrow(/<fetch>/);
        expect(() => buildQueryModel({ doc: parseFetchXml(`<fetch><entity name="unknowntable"/></fetch>`), tables, attributes, pageSize: 5000, includeRecordUrl: false })).toThrow(
            /unknowntable/,
        );
    });

    it("limits the preview FetchXML", () => {
        const { tables, attributes } = fixture();
        const model = buildQueryModel({ doc: parseFetchXml(fetch), tables, attributes, pageSize: 5000, includeRecordUrl: false });
        expect(previewFetchXml(model, 50)).toMatch(/^<fetch [^>]*top="50"/);
    });
});

describe("other queries", () => {
    it("generates valid parameter and choice queries", async () => {
        await expectValidM(generateParameterQuery(URL, DEFAULT_SETTINGS));
        expect(generateParameterQuery(URL, DEFAULT_SETTINGS)).toBe(
            '"https://contoso.crm4.dynamics.com/api/data/v9.2" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
        );
        expect(generateParameterQuery(URL, { ...DEFAULT_SETTINGS, apiVersion: "v9.1" })).toContain('"https://contoso.crm4.dynamics.com/api/data/v9.1"');
        const model = tableModel(["industrycode", "statecode", "statuscode", "donotemail", "new_tags", "name"]);
        const specs = choiceColumns(model.columns);
        expect(specs.map((s) => s.attribute.logicalName)).toEqual(["industrycode", "statecode", "statuscode", "donotemail", "new_tags"]);
        const { tables } = fixture();
        for (const spec of specs) {
            const live = generateChoiceQuery(spec, tables.get("account"), DEFAULT_SETTINGS, URL, null);
            await expectValidM(live.code);
            expect(live.code).toContain(`Attributes(LogicalName='${spec.attribute.logicalName}')`);
            expect(live.code).toContain("Web.Contents(DataverseUrl, [");
            expect(live.code).toContain(`RelativePath = "EntityDefinitions(LogicalName='account')`);
            const options =
                spec.attribute.kind === "boolean"
                    ? [
                          { value: true, label: 'Say "yes"' },
                          { value: false, label: "No" },
                      ]
                    : [
                          { value: 1, label: "Accounting #(x)", state: 0 },
                          { value: 2, label: "Line\nbreak", state: 1 },
                      ];
            const fixed = generateChoiceQuery(spec, tables.get("account"), DEFAULT_SETTINGS, URL, options);
            await expectValidM(fixed.code);
        }
        expect(generateChoiceQuery(specs[0], tables.get("account"), DEFAULT_SETTINGS, URL, null).queryName).toBe("Account Industry");
    });

    it("uses the organisation URL, not the Web API host", () => {
        expect(normaliseUrl("https://contoso.api.crm4.dynamics.com")).toBe("https://contoso.crm4.dynamics.com");
        expect(normaliseUrl("https://contoso.api.crm.dynamics.com/api/data/v9.2/")).toBe("https://contoso.crm.dynamics.com");
        expect(normaliseUrl("https://contoso.crm4.dynamics.com/")).toBe("https://contoso.crm4.dynamics.com");
        expect(normaliseUrl("https://contoso.api.crm.microsoftdynamics.us")).toBe("https://contoso.crm.microsoftdynamics.us");
        expect(normaliseUrl("https://api.example.com")).toBe("https://api.example.com");
        expect(generateParameterQuery("https://contoso.api.crm4.dynamics.com", DEFAULT_SETTINGS)).toContain(
            '"https://contoso.crm4.dynamics.com/api/data/v9.2"',
        );
        const settings = { ...DEFAULT_SETTINGS, useParameter: false, includeRecordUrl: true };
        const model = tableModel(["name"], settings);
        const code = generateFetchXmlQuery(model, resolveOutputColumns(model.columns, settings, {}), settings, "https://contoso.api.crm4.dynamics.com");
        expect(code).toContain('DataverseUrl = "https://contoso.crm4.dynamics.com/api/data/v9.2"');
        expect(code).not.toContain(".api.");
    });

    it("only sends headers Power Query allows with a credential", () => {
        const settings = { ...DEFAULT_SETTINGS, includeRecordUrl: true, lookupMode: "all" as const };
        const model = tableModel(["name", "ownerid", "industrycode", "donotemail"], settings);
        const out = resolveOutputColumns(model.columns, settings, {});
        const { tables } = fixture();
        const codes = [
            generateFetchXmlQuery(model, out, settings, URL),
            generateODataQuery(model, out, settings, URL),
            ...choiceColumns(model.columns).map((s) => generateChoiceQuery(s, tables.get("account"), settings, URL, null).code),
        ];
        for (const code of codes) {
            const headerBlocks = code.match(/Headers = \[[^\]]*\]/g) ?? [];
            expect(headerBlocks.length).toBeGreaterThan(0);
            for (const block of headerBlocks) {
                // Drop the quoted values first, so "odata.include-annotations=..." isn't read as a header name
                const withoutValues = block.replace(/=\s*"(?:[^"]|"")*"/g, "= _");
                const names = Array.from(withoutValues.matchAll(/(?:#"([^"]+)"|([A-Za-z-]+))\s*=/g))
                    .map((m) => m[1] ?? m[2])
                    .filter((n) => n !== "Headers");
                for (const n of names) expect(ALLOWED_AUTH_HEADERS).toContain(n);
            }
            expect(code).not.toContain("OData-");
        }
    });

    it("escapes M text", () => {
        expect(mText('a"b')).toBe('"a""b"');
        expect(mText("x#(lf)y")).toBe('"x#(#)(lf)y"');
        expect(mText("1\n2")).toBe('"1#(lf)2"');
    });
});

describe("preview", () => {
    it("applies the same mapping as the M query", () => {
        const settings = { ...DEFAULT_SETTINGS, includeRecordUrl: true };
        const model = tableModel(["name", "primarycontactid", "createdon", "new_startdate", "donotemail"], settings);
        const out = resolveOutputColumns(model.columns, settings, {});
        const preview = buildPreview(
            model,
            out,
            [
                {
                    accountid: "a1",
                    name: "Contoso",
                    _primarycontactid_value: "c1",
                    "_primarycontactid_value@OData.Community.Display.V1.FormattedValue": "Ada Lovelace",
                    createdon: "2024-03-05T21:59:01Z",
                    new_startdate: "2024-03-06",
                    donotemail: false,
                    extra: 1,
                },
            ],
            settings,
            URL,
        );
        expect(preview.columns).toEqual(["Account Name", "Primary Contact", "Primary Contact ID", "Created On", "Start Date", "Do not allow Emails", "Record URL"]);
        expect(preview.rows[0]).toEqual([
            "Contoso",
            "Ada Lovelace",
            "c1",
            "2024-03-05 21:59:01",
            "2024-03-06",
            "FALSE",
            "https://contoso.crm4.dynamics.com/main.aspx?pagetype=entityrecord&etn=account&id=a1",
        ]);
        expect(preview.unusedKeys).toEqual(["extra"]);
    });
});

describe("review fixes", () => {
    const build = (xml: string) => {
        const { tables, attributes } = fixture();
        return buildQueryModel({ doc: parseFetchXml(xml), tables, attributes, pageSize: 5000, includeRecordUrl: false });
    };

    it("keeps column comparisons pointing at a linked table whose generated alias is replaced", () => {
        const model = build(`<fetch><entity name="account"><attribute name="name" />
            <filter><condition attribute="name" operator="ne" valueof="a_0123456789abcdef0123456789abcdef.fullname" />
                    <condition entityname="a_0123456789abcdef0123456789abcdef" attribute="emailaddress1" operator="not-null" /></filter>
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="a_0123456789abcdef0123456789abcdef"><attribute name="fullname" /></link-entity>
        </entity></fetch>`);
        expect(model.fetchXml).toContain('valueof="contact.fullname"');
        expect(model.fetchXml).toContain('entityname="contact"');
        expect(model.fetchXml).not.toContain("a_0123456789abcdef");
    });

    it("warns about the paging limits Microsoft documents", () => {
        const sortedOnLink = build(`<fetch><entity name="account"><attribute name="name" />
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="fullname" /><order attribute="fullname" /></link-entity></entity></fetch>`);
        expect(sortedOnLink.warnings.join(" ")).toContain("stops at 50,000 rows");

        const distinct = build(`<fetch distinct="true"><entity name="account"><attribute name="name" /></entity></fetch>`);
        expect(distinct.warnings.join(" ")).toContain("distinct but has no <order>");

        // Fine: sorted on the main table; and a top query isn't paged at all
        expect(build(`<fetch distinct="true"><entity name="account"><attribute name="name" /><order attribute="name" /></entity></fetch>`).warnings).toEqual([]);
        const top = build(`<fetch top="10" distinct="true"><entity name="account"><attribute name="name" /></entity></fetch>`);
        expect(top.warnings.join(" ")).not.toContain("<order>");
    });
});

describe("unticked columns", () => {
    const build = (xml: string, excluded: string[], includeRecordUrl = false) => {
        const { tables, attributes } = fixture();
        return buildQueryModel({ doc: parseFetchXml(xml), tables, attributes, pageSize: 5000, includeRecordUrl, excluded: new Set(excluded) });
    };
    const XML = `<fetch><entity name="account"><attribute name="name" /><attribute name="entityimage_url" /><attribute name="revenue" />
        <link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="fullname" /><attribute name="emailaddress1" /></link-entity></entity></fetch>`;

    it("are taken out of the request, and warnings about them go with them", async () => {
        const all = build(XML, []);
        expect(all.warnings.join(" ")).toContain('"entityimage_url" wasn\'t found');

        const model = build(XML, ["|entityimage_url", "c|emailaddress1"]);
        expect(model.warnings).toEqual([]);
        expect(model.fetchXml).not.toContain("entityimage_url");
        expect(model.fetchXml).not.toContain("emailaddress1");
        expect(model.fetchXml).toContain('<attribute name="fullname" />');
        // The grid still lists them, so they can be ticked again
        expect(model.columns.map((c) => c.id)).toEqual(all.columns.map((c) => c.id));

        const settings = DEFAULT_SETTINGS;
        const outputs = resolveOutputColumns(model.columns, settings, { "|entityimage_url": { hidden: true }, "c|emailaddress1": { hidden: true } });
        const code = generateFetchXmlQuery(model, outputs, settings, "https://contoso.crm4.dynamics.com");
        expect(code).not.toContain("entityimage_url");
        const parsed = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, code);
        expect(PQP.TaskUtils.isParseStageOk(parsed)).toBe(true);
    });

    it("keeps the linked table's join when all its columns are unticked", () => {
        const model = build(XML, ["c|fullname", "c|emailaddress1"]);
        expect(model.fetchXml).toContain('<link-entity name="contact"');
        expect(model.fetchXml).not.toContain("fullname");
    });

    it("keeps the primary key in the request when the record URL needs it", () => {
        const xml = `<fetch><entity name="account"><attribute name="accountid" /><attribute name="name" /></entity></fetch>`;
        expect(build(xml, ["|accountid"]).fetchXml).not.toContain("accountid");
        expect(build(xml, ["|accountid"], true).fetchXml).toContain('<attribute name="accountid" />');
    });

    it("leaves requests alone where removing a column would change the rows", () => {
        // A distinct query can only sort on columns it returns
        const distinct = `<fetch distinct="true"><entity name="account"><attribute name="name" /><attribute name="revenue" /><order attribute="name" /></entity></fetch>`;
        const d = build(distinct, ["|name", "|revenue"]);
        expect(d.fetchXml).toContain('<attribute name="name" />');
        expect(d.fetchXml).not.toContain('<attribute name="revenue" />');
        // Aggregates: removing a group-by would change the grouping
        const aggregate = `<fetch aggregate="true"><entity name="account"><attribute name="industrycode" alias="industry" groupby="true" /><attribute name="revenue" alias="total" aggregate="sum" /></entity></fetch>`;
        expect(build(aggregate, ["|industry"]).fetchXml).toContain('alias="industry"');
    });

    it("says so when nothing is left", () => {
        const xml = `<fetch><entity name="account"><attribute name="name" /></entity></fetch>`;
        expect(build(xml, ["|name"]).warnings).toEqual(["The query doesn't return any columns yet."]);
    });
});
