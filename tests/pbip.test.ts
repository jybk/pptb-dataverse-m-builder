import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { describe, expect, it } from "vitest";
import { generateParameterQuery } from "../src/lib/mcode";
import { buildPowerBiProject, fileSafeName, type PowerBiProject } from "../src/lib/pbip";
import { URL, setup } from "./modelSetup";

const XML = `<fetch><entity name="account">
    <attribute name="accountid" /><attribute name="name" /><attribute name="industrycode" /><attribute name="new_tags" />
    <attribute name="primarycontactid" /><attribute name="ownerid" /><attribute name="createdby" /><attribute name="modifiedby" />
    <attribute name="revenue" /><attribute name="numberofemployees" /><attribute name="donotemail" /><attribute name="createdon" />
</entity></fetch>`;

function project(opts: Parameters<typeof setup>[1] = {}, name?: string): PowerBiProject {
    const { data, settings, mainCode } = setup(XML, { lookups: true, overrides: { "|industrycode": { mode: "both" } }, ...opts });
    return buildPowerBiProject({
        name: name ?? data.mainName,
        main: { name: data.mainName, code: mainCode, columns: data.mainColumns },
        queries: data.queries,
        relationships: data.relationships,
        parameter: settings.useParameter ? { name: settings.parameterName, code: generateParameterQuery(URL, settings) } : null,
    });
}

const file = (p: PowerBiProject, path: string) => {
    const f = p.files.find((x) => x.path === path);
    if (!f) throw new Error(`No ${path} in ${p.files.map((x) => x.path).join(", ")}`);
    return f.content;
};

/** PBIP_OUT=dir writes the projects to disk, for a check with Microsoft's TMDL reader. */
function writeOut(p: PowerBiProject, expected: Record<string, string> = {}) {
    const out = process.env.PBIP_OUT;
    if (!out) return;
    const write = (path: string, content: string) => {
        mkdirSync(dirname(path), { recursive: true });
        writeFileSync(path, content);
    };
    for (const f of p.files) write(join(out, p.name, f.path), f.content);
    // What each partition should read back as
    for (const [table, code] of Object.entries(expected)) write(join(out, `${p.name}.expected`, `${table}.m`), code.replace(/^\/\/ Suggested query name: .*\n/, ""));
}

function expectedM(opts: Parameters<typeof setup>[1] = {}): Record<string, string> {
    const { data, mainCode } = setup(XML, { lookups: true, overrides: { "|industrycode": { mode: "both" } }, ...opts });
    return Object.fromEntries([[data.mainName, mainCode], ...data.queries.map((q) => [q.queryName, q.code])]);
}

describe("Power BI project", () => {
    it("has a .pbip that opens the report, which opens the semantic model", () => {
        const p = project();
        writeOut(p, expectedM());
        expect(p.name).toBe("Account");
        expect(JSON.parse(file(p, "Account.pbip"))).toMatchObject({ version: "1.0", artifacts: [{ report: { path: "Account.Report" } }] });
        expect(JSON.parse(file(p, "Account.Report/definition.pbir")).datasetReference).toEqual({ byPath: { path: "../Account.SemanticModel" } });
        expect(JSON.parse(file(p, "Account.SemanticModel/definition.pbism")).version).toBe("4.0");
        const report = JSON.parse(file(p, "Account.Report/report.json"));
        expect(report.sections).toHaveLength(1);
        expect(JSON.parse(report.config).themeCollection.baseTheme.name).toBe("CY24SU10");
        expect(JSON.parse(file(p, "Account.Report/StaticResources/SharedResources/BaseThemes/CY24SU10.json")).name).toBe("CY24SU10");
        expect(file(p, ".gitignore")).toContain("**/.pbi/cache.abf");
    });

    it("has a table per query, the parameter first in the query order", () => {
        const p = project();
        const model = file(p, "Account.SemanticModel/definition/model.tmdl");
        expect(model).toContain(
            'annotation PBI_QueryOrder = ["DataverseUrl","Account","Account Do not allow Emails","Account Industry","Account Tags","Account Tags bridge","Contact","Owner","User"]',
        );
        expect(model).toContain("ref table 'Account Tags bridge'\nref table Contact\n");
        const tables = p.files.filter((f) => f.path.includes("/tables/")).map((f) => f.path.split("/").pop());
        expect(tables).toEqual(["Account.tmdl", "Account Do not allow Emails.tmdl", "Account Industry.tmdl", "Account Tags.tmdl", "Account Tags bridge.tmdl", "Contact.tmdl", "Owner.tmdl", "User.tmdl"]);
        expect(file(p, "Account.SemanticModel/definition/expressions.tmdl")).toMatch(
            /^expression DataverseUrl = "https:\/\/contoso\.crm4\.dynamics\.com\/api\/data\/v9\.2" meta \[IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true\]\n\tlineageTag: [0-9a-f-]{36}\n\n\tannotation PBI_ResultType = Text\n$/,
        );
        const yesNo = file(p, "Account.SemanticModel/definition/tables/Account Do not allow Emails.tmdl");
        expect(yesNo).toContain("\tcolumn Value\n\t\tdataType: boolean\n");
    });

    it("declares the main table's columns with their types, and the query as its partition", () => {
        const t = file(project({ recordUrl: true, settings: { includeRecordUrl: true } }), "Account.SemanticModel/definition/tables/Account.tmdl");
        const column = (name: string) => t.split("\n\tcolumn ").find((c) => c.startsWith(name.includes(" ") ? `'${name}'` : name))!;
        expect(column("Annual Revenue")).toContain("dataType: decimal");
        expect(column("Annual Revenue")).toContain("summarizeBy: sum");
        expect(column("Number of Employees")).toContain("dataType: int64\n\t\tformatString: 0");
        expect(column("Number of Employees")).toContain("summarizeBy: sum");
        expect(column("Industry Value")).toContain("dataType: int64");
        expect(column("Industry Value")).toContain("summarizeBy: none");
        expect(column("Created On")).toContain("dataType: dateTime\n\t\tformatString: General Date");
        expect(column("Do not allow Emails")).toContain("dataType: boolean");
        expect(column("Record URL")).toContain("dataCategory: WebUrl");
        expect(column("Owner ID")).toContain("sourceColumn: Owner ID");
        expect(t).toContain("\tpartition Account = m\n\t\tmode: import\n\t\tsource =\n\t\t\t\t// Generated by Dataverse M Builder");
        expect(t).not.toContain("Suggested query name");
        // Every line of the query is inside the partition, blank lines too
        const source = t.slice(t.indexOf("\t\tsource =\n") + 11, t.indexOf("\n\n\tannotation PBI_ResultType"));
        expect(source.split("\n").every((l) => l.startsWith("\t\t\t\t"))).toBe(true);
    });

    it("hides bridge tables and lookup keys, and writes the relationships", () => {
        const p = project();
        expect(file(p, "Account.SemanticModel/definition/tables/Account Tags bridge.tmdl")).toMatch(/^table 'Account Tags bridge'\n\tisHidden\n/);
        const owner = file(p, "Account.SemanticModel/definition/tables/Owner.tmdl");
        expect(owner).toContain("\tcolumn 'Owner key'\n\t\tdataType: string\n\t\tisHidden\n");
        expect(owner).toContain("\tcolumn 'Owner type'\n\t\tdataType: string\n\t\tlineageTag:");
        const rel = file(p, "Account.SemanticModel/definition/relationships.tmdl");
        expect(rel).toMatch(/^relationship [0-9a-f-]{36}\n\tfromColumn: /);
        expect(rel).toContain("\tisActive: false\n\tfromColumn: Account.'Modified By ID'\n\ttoColumn: User.'User key'");
        expect(rel).toContain("\tcrossFilteringBehavior: bothDirections\n\tfromColumn: 'Account Tags bridge'.'Account Tags bridge key'\n\ttoColumn: Account.'Account ID'");
    });

    it("gives the same ids on every export, so a project in Git only changes where the model does", () => {
        const a = project();
        const b = project();
        expect(b.files).toEqual(a.files);
    });

    it("leaves out the parameter when the URL is written into the queries", () => {
        const p = project({ settings: { useParameter: false } }, "Written URL");
        writeOut(p, expectedM({ settings: { useParameter: false } }));
        expect(p.files.some((f) => f.path.endsWith("expressions.tmdl"))).toBe(false);
        expect(file(p, "Written URL.SemanticModel/definition/model.tmdl")).toContain('PBI_QueryOrder = ["Account",');
    });

    it("handles names that need quoting or encoding", () => {
        const overrides = {
            "|industrycode": { mode: "both" as const },
            "|name": { name: '"Quoted" name' },
            "|revenue": { name: "Revenue: 'net'" },
            "|createdon": { name: "Created.On #1" },
        };
        const p = project({ mainName: 'Klant\'s "beste"', overrides }, 'Klant\'s "beste": export.');
        writeOut(p, expectedM({ mainName: 'Klant\'s "beste"', overrides }));
        expect(p.name).toBe("Klant's beste export");
        const main = file(p, "Klant's beste export.SemanticModel/definition/tables/Klants beste.tmdl");
        expect(main).toContain(`\tcolumn '"Quoted" name'\n`);
        expect(main).toContain(`\t\tsourceColumn: """Quoted"" name"\n`);
        expect(main).toContain(`\tcolumn 'Revenue: ''net'''\n`);
        expect(main).toContain(`\t\tsourceColumn: Revenue: 'net'\n`);
        expect(fileSafeName("  ")).toBe("Dataverse");
        expect(fileSafeName("a/b\\c:d*e?f")).toBe("abcdef");
    });

    it("declares a lookup table's added columns", () => {
        const lookupColumns = { contact: { columns: ["emailaddress1", "parentcustomerid"], overrides: {} } };
        const p = project({ lookupColumns }, "Extended");
        writeOut(p, expectedM({ lookupColumns }));
        const contact = file(p, "Extended.SemanticModel/definition/tables/Contact.tmdl");
        expect(contact).toContain("\tcolumn 'Contact key'\n\t\tdataType: string\n\t\tisHidden\n");
        expect(contact).toContain("\tcolumn Email\n\t\tdataType: string\n");
        expect(contact).toContain("\tcolumn 'Company Name ID'\n");
        expect(contact).toContain('#"$select" = "contactid,fullname,emailaddress1,_parentcustomerid_value"');
    });
});
