import * as PQP from "@microsoft/powerquery-parser";
import { resolveOutputColumns } from "../src/lib/columns";
import { buildQueryModel, parseFetchXml, type ModelInput } from "../src/lib/fetchxml";
import { choiceColumns, generateChoiceQuery, generateFetchXmlQuery } from "../src/lib/mcode";
import { buildDataModel } from "../src/lib/model";
import { mapAttributes } from "../src/lib/metadata";
import { DEFAULT_SETTINGS, type AttributeInfo, type BuilderSettings, type ColumnOverrides, type LookupColumns, type TableInfo } from "../src/lib/types";
import { attr, fixture, rawAccountAttributes, rawAccountDates, rawAccountLookups, rawContactAttributes } from "./fixtures";

/** Shared by the model and Power BI project tests: a query and everything the Model tab builds from it. */
export const URL = "https://contoso.crm4.dynamics.com";

export async function expectValidM(code: string) {
    const result = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, code);
    if (!PQP.TaskUtils.isParseStageOk(result)) throw new Error(`Generated M doesn't parse:\n\n${code}`);
}

export interface SetupOptions {
    overrides?: ColumnOverrides;
    canAddColumns?: boolean;
    mainName?: string;
    settings?: Partial<BuilderSettings>;
    /** Lookups and tables beyond the shared fixture. */
    lookups?: boolean;
    recordUrl?: boolean;
    lookupColumns?: Record<string, LookupColumns>;
}

export const LOOKUP_TABLES: TableInfo[] = [
    { logicalName: "systemuser", displayName: "User", entitySetName: "systemusers", primaryIdAttribute: "systemuserid", primaryNameAttribute: "fullname", isCustom: false, isIntersect: false, isActivity: false },
    { logicalName: "team", displayName: "Team", entitySetName: "teams", primaryIdAttribute: "teamid", primaryNameAttribute: "name", isCustom: false, isIntersect: false, isActivity: false },
    { logicalName: "new_gadget", displayName: "Gadget", entitySetName: "new_gadgets", primaryIdAttribute: "new_gadgetid", primaryNameAttribute: null, isCustom: true, isIntersect: false, isActivity: false },
];

export function withLookups(tables: Map<string, TableInfo>, attributes: Map<string, AttributeInfo[]>) {
    for (const t of LOOKUP_TABLES) tables.set(t.logicalName, t);
    tables.delete("new_gadget");
    const extra = [
        attr("createdby", "Created By", "LookupType"),
        attr("modifiedby", "Modified By", "LookupType"),
        attr("parentaccountid", "Parent Account", "LookupType"),
        attr("new_approverid", "Approver", "OwnerType"),
        attr("new_regardingid", "Regarding", "LookupType"),
        attr("new_mysteryid", "Mystery", "LookupType"),
        attr("new_gadgetid", "Gadget", "LookupType"),
        attr("new_deviceid", "Device", "LookupType"),
    ];
    const lookups = [
        ...rawAccountLookups,
        { LogicalName: "createdby", Targets: ["systemuser"] },
        { LogicalName: "modifiedby", Targets: ["systemuser"] },
        { LogicalName: "parentaccountid", Targets: ["account"] },
        { LogicalName: "new_approverid", Targets: ["team", "systemuser"] },
        { LogicalName: "new_regardingid", Targets: ["account", "contact", "new_project", "systemuser"] },
        { LogicalName: "new_gadgetid", Targets: ["new_gadget"] },
        { LogicalName: "new_deviceid", Targets: ["new_project"] },
    ];
    attributes.set("account", mapAttributes([...rawAccountAttributes, ...extra], { dateTimes: rawAccountDates, lookups }));
    attributes.set("contact", mapAttributes(rawContactAttributes, { lookups: [{ LogicalName: "parentcustomerid", Targets: ["contact", "account"] }] }));
}

export function setup(xml: string, opts: SetupOptions = {}) {
    const { tables, attributes } = fixture();
    if (opts.lookups) withLookups(tables, attributes);
    const settings = { ...DEFAULT_SETTINGS, ...opts.settings };
    const overrides = opts.overrides ?? {};
    const excluded = new Set(
        Object.entries(overrides)
            .filter(([, o]) => o.hidden)
            .map(([id]) => id),
    );
    const modelInput: ModelInput = { doc: parseFetchXml(xml), tables, attributes, pageSize: 5000, includeRecordUrl: !!opts.recordUrl, excluded };
    const model = buildQueryModel(modelInput);
    const outputs = resolveOutputColumns(model.columns, settings, overrides, !!model.recordUrlKey);
    const choiceQueries = choiceColumns(model.columns.filter((c) => !overrides[c.id]?.hidden)).map((spec) =>
        generateChoiceQuery(spec, tables.get(spec.entity), settings, URL, null),
    );
    const data = buildDataModel({
        model,
        outputs,
        mainName: opts.mainName ?? "",
        choiceQueries,
        modelInput,
        settings,
        connectionUrl: URL,
        canAddColumns: !!opts.canAddColumns,
        lookupColumns: opts.lookupColumns,
    });
    const mainCode = generateFetchXmlQuery(model, outputs, settings, URL);
    return { model, outputs, data, settings, mainCode, tables };
}
