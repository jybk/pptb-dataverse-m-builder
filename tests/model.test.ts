import { describe, expect, it } from "vitest";
import { resolveOutputColumns } from "../src/lib/columns";
import { buildQueryModel, buildTableFetchXml, parseFetchXml, type ModelInput } from "../src/lib/fetchxml";
import { buildDataModel, stableId, tmdlName, tmdlScript } from "../src/lib/model";
import { DEFAULT_SETTINGS } from "../src/lib/types";
import { fixture } from "./fixtures";
import { URL, expectValidM, setup, withLookups } from "./modelSetup";




const TABLE_XML = () => {
    const { tables } = fixture();
    return buildTableFetchXml(tables.get("account")!, ["accountid", "name", "industrycode", "new_tags", "donotemail"], true);
};

describe("multi-select bridge tables", () => {
    it("are built from a light request with the same filters, and split the codes into rows", async () => {
        const { data } = setup(TABLE_XML(), { overrides: { "|industrycode": { mode: "both" } } });
        const bridge = data.queries.find((q) => q.kind === "bridge")!;
        expect(bridge.queryName).toBe("Account Tags bridge");
        await expectValidM(bridge.code);
        expect(bridge.code).toContain("// Suggested query name: Account Tags bridge");
        // Only the key and the multi-select column are requested, under the main query's filter
        expect(bridge.code).toContain('"<attribute name=""accountid"" />"');
        expect(bridge.code).toContain('"<attribute name=""new_tags"" />"');
        expect(bridge.code).not.toContain('name=""name""');
        expect(bridge.code).not.toContain("industrycode");
        expect(bridge.code).toContain('"<condition attribute=""statecode"" operator=""eq"" value=""0"" />"');
        expect(bridge.code).toContain('Text.Split([Codes], ",")');
        expect(bridge.code).toContain('Table.TransformColumnTypes(Selected, {{"Account Tags bridge value", Int64.Type}})');
        expect(bridge.code).toMatch(/in\n {4}Bridge$/);
        // Paging works the same as in the main query
        expect(bridge.code).toContain("PagingCookie = NextCookie");
    });

    it("sit right after their choice table in the list", () => {
        const { data } = setup(TABLE_XML());
        expect(data.queries.map((q) => `${q.kind}:${q.queryName}`)).toEqual([
            "choice:Account Do not allow Emails",
            "choice:Account Industry",
            "choice:Account Tags",
            "bridge:Account Tags bridge",
        ]);
    });

    it("are left out where the query can't return the primary key", () => {
        const { data } = setup(`<fetch distinct="true"><entity name="account"><attribute name="new_tags" /><order attribute="new_tags" /></entity></fetch>`);
        expect(data.queries.some((q) => q.kind === "bridge")).toBe(false);
        expect(data.notes.map((n) => n.text).join(" ")).toContain("aggregate and distinct queries don't return the primary key");
    });

    it("are only made for the main table's own columns", () => {
        const xml = `<fetch><entity name="contact"><attribute name="contactid" /><attribute name="fullname" />
            <link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><attribute name="new_tags" /></link-entity></entity></fetch>`;
        const { data } = setup(xml);
        expect(data.queries.some((q) => q.kind === "bridge")).toBe(false);
        expect(data.notes.map((n) => n.text).join(" ")).toContain("multi-select column of a linked table");
    });
});

describe("relationships", () => {
    it("relate choice tables to value columns, and bridges both ways to the main query", () => {
        const { data, outputs } = setup(TABLE_XML(), { overrides: { "|industrycode": { mode: "both" } } });
        const industryValue = outputs.find((o) => o.colId === "|industrycode" && o.role === "value")!.name;
        const pk = outputs.find((o) => o.colId === "|accountid")!.name;
        const list = data.relationships.map(
            (r) => `${r.from.table}[${r.from.column}] -> ${r.to.table}[${r.to.column}]${r.bothDirections ? " both" : ""}${r.active ? "" : " inactive"}`,
        );
        expect(list).toEqual([
            `Account[${industryValue}] -> Account Industry[Value]`,
            "Account Tags bridge[Account Tags bridge value] -> Account Tags[Value]",
            `Account Tags bridge[Account Tags bridge key] -> Account[${pk}] both`,
        ]);
    });

    it("offer a fix when a choice only returns its label", () => {
        const { data } = setup(TABLE_XML());
        const note = data.notes.find((n) => n.text.startsWith("Industry only returns its label"))!;
        expect(note.fix).toEqual({ kind: "mode", colId: "|industrycode", mode: "both", label: "Return label + value" });
        expect(data.relationships.some((r) => r.to.table === "Account Industry")).toBe(false);
        expect(data.notes.map((n) => n.text).join(" ")).toContain("Yes/no columns aren't related");
    });

    it("ask for the primary key when the bridge has nothing to relate to", () => {
        const { tables } = fixture();
        const withoutKey = buildTableFetchXml(tables.get("account")!, ["name", "new_tags"], false);
        const table = setup(withoutKey, { canAddColumns: true });
        expect(table.data.notes.find((n) => n.text.startsWith("Add the primary key"))?.fix).toEqual({
            kind: "include",
            colId: "|accountid",
            label: "Add the primary key",
        });

        // FetchXML without the key: it has to be added to the FetchXML itself
        const fetch = setup(`<fetch><entity name="account"><attribute name="name" /><attribute name="new_tags" /></entity></fetch>`);
        const note = fetch.data.notes.find((n) => n.text.startsWith("Add the primary key"))!;
        expect(note.fix).toBeUndefined();
        expect(note.text).toContain("Add it to the FetchXML first");

        // FetchXML with the key unticked: ticking it again is enough
        const unticked = setup(`<fetch><entity name="account"><attribute name="accountid" /><attribute name="new_tags" /></entity></fetch>`, {
            overrides: { "|accountid": { hidden: true } },
        });
        expect(unticked.data.notes.find((n) => n.text.startsWith("Add the primary key"))?.fix).toMatchObject({ kind: "include", colId: "|accountid" });
        // The bridge still gets the key: it asks for it itself
        expect(unticked.data.queries.find((q) => q.kind === "bridge")!.code).toContain('"<attribute name=""accountid"" />"');
    });

    it("keep one active relationship per pair of tables", () => {
        const xml = `<fetch><entity name="account"><attribute name="name" />
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c1"><attribute name="gendercode" /></link-entity>
            <link-entity name="contact" from="contactid" to="primarycontactid" alias="c2"><attribute name="gendercode" /></link-entity></entity></fetch>`;
        const { data } = setup(xml, { overrides: { "c1|gendercode": { mode: "both" }, "c2|gendercode": { mode: "both" } } });
        expect(data.queries.filter((q) => q.kind === "choice")).toHaveLength(1);
        expect(data.relationships.map((r) => r.active)).toEqual([true, false]);
        expect(data.tmdl).toContain("\t\tisActive: false");
        const inactive = data.relationships[1].from.column;
        expect(data.notes.map((n) => n.text)).toContain(
            `${inactive} relates to Contact Gender through an inactive relationship: Power BI allows one active relationship between two tables. Use USERELATIONSHIP in a measure to filter by it.`,
        );
    });

    it("use the main query name you give, cleaned up", () => {
        const { data } = setup(TABLE_XML(), { mainName: " Klanten: actief ", overrides: { "|industrycode": { mode: "both" } } });
        expect(data.mainName).toBe("Klanten actief");
        expect(data.relationships[0].from.table).toBe("Klanten actief");
    });
});

const rel = (r: { from: { table: string; column: string }; to: { table: string; column: string }; bothDirections: boolean; active: boolean }) =>
    `${r.from.table}[${r.from.column}] -> ${r.to.table}[${r.to.column}]${r.bothDirections ? " both" : ""}${r.active ? "" : " inactive"}`;

const fetch = (attrs: string[], extra = "") =>
    `<fetch><entity name="account">${attrs.map((a) => `<attribute name="${a}" />`).join("")}${extra}</entity></fetch>`;

describe("lookup tables", () => {
    it("are made per target table, with key and name, reading every row", async () => {
        const { data } = setup(fetch(["accountid", "name", "primarycontactid"]), { lookups: true });
        const contact = data.queries.find((q) => q.kind === "lookup")!;
        expect(contact.queryName).toBe("Contact");
        expect(contact.source).toBe("contact for Primary Contact");
        await expectValidM(contact.code);
        expect(contact.code).toContain("// Suggested query name: Contact");
        expect(contact.code).toContain('Result = Rows("contacts", "contactid", "fullname")');
        expect(contact.code).toContain('#"Contact key" = Record.FieldOrDefault(r, keyField),');
        expect(contact.code).toContain('type table [#"Contact key" = nullable text, #"Contact name" = nullable text]');
        expect(contact.code).toContain('GetPage(Uri.Parts(Record.Field(_, "@odata.nextLink"))[Query])');
        expect(contact.code).toContain('Prefer = "odata.maxpagesize=5000"');
        // No filter: inactive rows are needed too, or records pointing to them find no match
        expect(contact.code).not.toContain("$filter");
        expect(contact.code).toMatch(/in\n {4}Result$/);
        expect(data.relationships.map(rel)).toEqual(["Account[Primary Contact ID] -> Contact[Contact key]"]);
    });

    it("combine user and team into Owner, and account and contact into Customer, with a type column", async () => {
        const xml = fetch(["name", "ownerid"], `<link-entity name="contact" from="contactid" to="primarycontactid" alias="pc"><attribute name="parentcustomerid" /></link-entity>`);
        const { data } = setup(xml, { lookups: true });
        const owner = data.queries.find((q) => q.queryName === "Owner")!;
        const customer = data.queries.find((q) => q.queryName === "Customer")!;
        for (const q of [owner, customer]) await expectValidM(q.code);
        expect(owner.source).toBe("systemuser + team for Owner");
        expect(owner.code).toContain('#"systemuser rows" = Table.AddColumn(Rows("systemusers", "systemuserid", "fullname"), "Owner type", each "User", type text),');
        expect(owner.code).toContain('#"team rows" = Table.AddColumn(Rows("teams", "teamid", "name"), "Owner type", each "Team", type text),');
        expect(owner.code).toContain('Result = Table.Combine({#"systemuser rows", #"team rows"})');
        expect(customer.code).toContain('Rows("accounts", "accountid", "name"), "Customer type", each "Account", type text)');
        // A linked table's lookup relates the same way
        const fromCustomer = data.relationships.find((r) => r.to.table === "Customer")!;
        expect(fromCustomer.to.column).toBe("Customer key");
        expect(fromCustomer.active).toBe(true);
    });

    it("are shared by lookups to the same tables: the first relationship active, Owner first", () => {
        const { data } = setup(fetch(["name", "createdby", "modifiedby", "new_approverid", "ownerid"]), { lookups: true });
        expect(data.queries.filter((q) => q.kind === "lookup").map((q) => q.queryName)).toEqual(["User", "Owner"]);
        expect(data.relationships.map(rel)).toEqual([
            "Account[Created By ID] -> User[User key]",
            "Account[Modified By ID] -> User[User key] inactive",
            "Account[Owner ID] -> Owner[Owner key]",
            "Account[Approver ID] -> Owner[Owner key] inactive",
        ]);
        const texts = data.notes.map((n) => n.text);
        expect(texts).toContain(
            "Modified By ID relates to User through an inactive relationship: Power BI allows one active relationship between two tables. Use USERELATIONSHIP in a measure to filter by it.",
        );
        expect(texts.some((t) => t.startsWith("Approver ID relates to Owner"))).toBe(true);
        expect(data.tmdl).toContain("\t\tisActive: false\n\t\tfromColumn: Account.'Modified By ID'\n\t\ttoColumn: User.'User key'");
    });

    it("get a table of their own for a lookup to the main table, named after the column", () => {
        const { data } = setup(fetch(["name", "parentaccountid"]), { lookups: true });
        expect(data.queries.find((q) => q.kind === "lookup")!.queryName).toBe("Parent Account");
        expect(data.relationships.map(rel)).toEqual(["Account[Parent Account ID] -> Parent Account[Parent Account key]"]);
    });

    it("don't take a name that's in use", () => {
        const { data } = setup(fetch(["name", "createdby", "modifiedby"]), { lookups: true, mainName: "user" });
        expect(data.queries.find((q) => q.kind === "lookup")!.queryName).toBe("User lookup");
        const single = setup(fetch(["name", "createdby"]), { lookups: true, mainName: "User" });
        expect(single.data.queries.find((q) => q.kind === "lookup")!.queryName).toBe("Created By");
    });

    it("offer a fix when a lookup only returns its name", () => {
        const { data } = setup(fetch(["name", "primarycontactid"]), { lookups: true, overrides: { "|primarycontactid": { mode: "name" } } });
        expect(data.queries.some((q) => q.kind === "lookup")).toBe(false);
        expect(data.notes.find((n) => n.text.startsWith("Primary Contact only returns its name"))?.fix).toEqual({
            kind: "mode",
            colId: "|primarycontactid",
            mode: "both",
            label: "Return name + ID",
        });
    });

    it("are left out, with a note, for many targets, unknown targets and tables that aren't listed", () => {
        const { data } = setup(fetch(["name", "new_regardingid", "new_mysteryid", "new_gadgetid"]), { lookups: true });
        expect(data.queries.some((q) => q.kind === "lookup")).toBe(false);
        expect(data.relationships).toEqual([]);
        expect(data.notes.map((n) => n.text)).toEqual([
            "Regarding can point to 4 tables, so it has no lookup table. Its name in the main query is what a report can use.",
            "Mystery: the tables it points to aren't known, so it has no lookup table.",
            "Gadget points to new_gadget, which isn't in the environment's table list, so it has no lookup table.",
        ]);
    });

    it("use an empty name column for a table without a primary name", async () => {
        const { tables, attributes } = fixture();
        withLookups(tables, attributes);
        const project = { ...tables.get("new_project")!, primaryNameAttribute: null };
        tables.set("new_project", project);
        const modelInput: ModelInput = { doc: parseFetchXml(fetch(["name", "new_deviceid"])), tables, attributes, pageSize: 5000, includeRecordUrl: false };
        const model = buildQueryModel(modelInput);
        const outputs = resolveOutputColumns(model.columns, DEFAULT_SETTINGS, {}, false);
        const data = buildDataModel({ model, outputs, mainName: "", choiceQueries: [], modelInput, settings: DEFAULT_SETTINGS, connectionUrl: URL, canAddColumns: false });
        const q = data.queries.find((x) => x.kind === "lookup")!;
        await expectValidM(q.code);
        expect(q.code).toContain('Result = Rows("new_projects", "new_projectid", null)');
    });

    it("can be switched off", () => {
        const { data } = setup(fetch(["name", "primarycontactid", "ownerid", "new_regardingid"]), { lookups: true, settings: { lookupTables: false } });
        expect(data.queries).toEqual([]);
        expect(data.notes).toEqual([]);
        expect(data.tmdl).toBeNull();
    });

    it("work with the URL written into the query", async () => {
        const { data } = setup(fetch(["name", "ownerid"]), { lookups: true, settings: { useParameter: false } });
        const owner = data.queries.find((q) => q.kind === "lookup")!;
        await expectValidM(owner.code);
        expect(owner.code).toContain('DataverseUrl = "https://contoso.crm4.dynamics.com/api/data/v9.2",');
        expect(owner.code).not.toContain("Needs a text parameter");
    });
});

describe("lookup tables with more columns", () => {
    const contactColumns = { contact: { columns: ["emailaddress1", "parentcustomerid", "nonexistent"], overrides: { "|emailaddress1": { name: "Contact e-mail" } } } };

    it("read the target table with the added columns, key and name under the lookup table's names", async () => {
        const { data } = setup(fetch(["name", "primarycontactid"]), { lookups: true, lookupColumns: contactColumns });
        const contact = data.queries.find((q) => q.kind === "lookup")!;
        await expectValidM(contact.code);
        expect(contact.code).toMatch(/^\/\/ Suggested query name: Contact\n/);
        expect(contact.code).toContain("// Lookup table: Contact (contact) for Primary Contact. Every row is read, active or not");
        expect(contact.code).toContain('#"$select" = "contactid,fullname,emailaddress1,_parentcustomerid_value"');
        expect(contact.code).not.toContain("$filter");
        expect(contact.columns.map((c) => c.name)).toEqual(["Contact key", "Contact name", "Contact e-mail", "Company Name", "Company Name ID"]);
        expect(contact.columns[0].hidden).toBe(true);
        expect(contact.addedColumns).toBe(2);
        expect(contact.target?.logicalName).toBe("contact");
        expect(data.relationships.map(rel)).toEqual(["Account[Primary Contact ID] -> Contact[Contact key]"]);
    });

    it("stay key and name when nothing valid is added", () => {
        const { data } = setup(fetch(["name", "primarycontactid"]), { lookups: true, lookupColumns: { contact: { columns: ["contactid", "nonexistent"], overrides: {} } } });
        const contact = data.queries.find((q) => q.kind === "lookup")!;
        expect(contact.code).toContain('Result = Rows("contacts", "contactid", "fullname")');
        expect(contact.addedColumns).toBeUndefined();
    });

    it("apply to lookup tables of one table only, not to combined ones", () => {
        const { data } = setup(fetch(["name", "ownerid", "createdby"]), {
            lookups: true,
            lookupColumns: { systemuser: { columns: ["fullname"], overrides: {} }, team: { columns: ["name"], overrides: {} } },
        });
        const owner = data.queries.find((q) => q.queryName === "Owner")!;
        expect(owner.target).toBeUndefined();
        expect(owner.code).toContain("Table.Combine");
    });
});

describe("TMDL script", () => {
    it("quotes names the way TMDL needs", () => {
        expect(tmdlName("Account")).toBe("Account");
        expect(tmdlName("Account_2")).toBe("Account_2");
        expect(tmdlName("Account Tags")).toBe("'Account Tags'");
        expect(tmdlName("O'Brien")).toBe("'O''Brien'");
        expect(tmdlName("a.b")).toBe("'a.b'");
        expect(tmdlName("Größe")).toBe("'Größe'");
        expect(tmdlName("2025")).toBe("'2025'");
    });

    it("gives each relationship a stable GUID", () => {
        const id = stableId("Account[Industry] > Account Industry[Value]");
        expect(id).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
        expect(stableId("Account[Industry] > Account Industry[Value]")).toBe(id);
        expect(stableId("Account[Industry] > Account Industry[Label]")).not.toBe(id);
    });

    it("writes one createOrReplace script with tab indentation", () => {
        const { data } = setup(TABLE_XML(), { overrides: { "|industrycode": { mode: "both" } } });
        const lines = data.tmdl!.split("\n");
        expect(lines[0]).toBe("createOrReplace");
        expect(lines.filter((l) => l.startsWith("\trelationship "))).toHaveLength(3);
        expect(data.tmdl).toContain("\t\tfromColumn: 'Account Tags bridge'.'Account Tags bridge value'\n\t\ttoColumn: 'Account Tags'.Value");
        expect(data.tmdl).toContain("\t\tcrossFilteringBehavior: bothDirections");
        expect(data.tmdl).not.toContain("isActive");
        expect(tmdlScript([])).toBeNull();
    });
});
