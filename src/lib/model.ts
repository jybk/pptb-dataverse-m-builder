import { resolveOutputColumns } from "./columns";
import { buildQueryModel, buildTableFetchXml, parseFetchXml, type ModelInput } from "./fetchxml";
import { bridgeColumns, bridgeQueryName, generateBridgeQuery, generateLookupQuery, generateODataQuery, lookupColumns, type ChoiceQuery, type QueryColumn } from "./mcode";
import { recordUrlColumnName } from "./columns";
import { isChoiceLike } from "./metadata";
import type { BuilderSettings, ColumnOverrides, ColumnSpec, LookupColumns, OutputColumn, QueryModel, TableInfo } from "./types";

/** One query of the model besides the main one. */
export interface ModelQuery {
    id: string;
    kind: "choice" | "bridge" | "lookup";
    queryName: string;
    /** What it's made from, e.g. "account.new_colors", or for a lookup table "systemuser + team for Owner". */
    source: string;
    code: string;
    /** What it returns, for exports that declare the table's columns. */
    columns: QueryColumn[];
    /** For a lookup table of one table: that table, whose other columns can be added. */
    target?: TableInfo;
    /** Lookup tables: how many columns were added besides key and name. */
    addedColumns?: number;
}

export interface Relationship {
    /** Stable for the same pair of columns, so applying the script again replaces it instead of adding a copy. */
    id: string;
    /** The many side. */
    from: { table: string; column: string };
    /** The one side. */
    to: { table: string; column: string };
    bothDirections: boolean;
    active: boolean;
}

export interface ModelNote {
    text: string;
    /** A change in the main query that solves it: tick a column, or change a column's output. */
    fix?: { kind: "include"; colId: string; label: string } | { kind: "mode"; colId: string; mode: string; label: string };
}

export interface DataModel {
    mainName: string;
    /** The main query's columns, in order, for exports that declare them. */
    mainColumns: QueryColumn[];
    queries: ModelQuery[];
    relationships: Relationship[];
    notes: ModelNote[];
    /** createOrReplace script for TMDL view, or null without relationships. */
    tmdl: string | null;
}

/** A query name Power BI accepts as is: letters, digits, spaces, _ and -. */
export function cleanQueryName(name: string): string {
    return name
        .replace(/[^\p{L}\p{N} _-]/gu, "")
        .replace(/\s+/g, " ")
        .trim();
}

export function suggestedMainName(table: TableInfo): string {
    return cleanQueryName(table.displayName) || table.logicalName;
}

/** TMDL needs single quotes around names with anything but letters, digits and _, and doubles a quote inside. */
export function tmdlName(name: string): string {
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** A GUID-shaped id derived from text (cyrb128), so the same relationship always gets the same id. */
export function stableId(text: string): string {
    let h1 = 1779033703,
        h2 = 3144134277,
        h3 = 1013904242,
        h4 = 2773480762;
    for (let i = 0; i < text.length; i++) {
        const k = text.charCodeAt(i);
        h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
        h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
        h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
        h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
    }
    h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
    h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
    h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
    h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
    const hex = [h1 ^ h2 ^ h3 ^ h4, h2 ^ h1, h3 ^ h1, h4 ^ h1].map((h) => (h >>> 0).toString(16).padStart(8, "0")).join("");
    // Version 4 and variant bits, so it reads as an ordinary GUID
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-${((parseInt(hex[16], 16) & 3) | 8).toString(16)}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

/** A relationship in TMDL, starting at the given indent (one tab per level). */
export function relationshipTmdl(r: Relationship, depth = 0): string {
    const at = "\t".repeat(depth);
    return [
        `${at}relationship ${r.id}`,
        ...(r.bothDirections ? [`${at}\tcrossFilteringBehavior: bothDirections`] : []),
        ...(r.active ? [] : [`${at}\tisActive: false`]),
        `${at}\tfromColumn: ${tmdlName(r.from.table)}.${tmdlName(r.from.column)}`,
        `${at}\ttoColumn: ${tmdlName(r.to.table)}.${tmdlName(r.to.column)}`,
    ].join("\n");
}

export function tmdlScript(relationships: Relationship[]): string | null {
    if (relationships.length === 0) return null;
    return ["createOrReplace", "", relationships.map((r) => relationshipTmdl(r, 1)).join("\n\n"), ""].join("\n");
}

const NUMBER_KINDS = new Set(["integer", "bigint", "decimal", "double", "money"]);

/** The main query's columns as it returns them, the Record URL last. */
function mainQueryColumns({ model, outputs, settings }: DataModelInput): QueryColumn[] {
    return queryColumns(model, outputs, settings);
}

/** A query's columns as it returns them, for exports: the output columns, then the Record URL if it has one. */
function queryColumns(model: QueryModel, outputs: OutputColumn[], settings: BuilderSettings): QueryColumn[] {
    const specs = new Map(model.columns.map((c) => [c.id, c]));
    const columns: QueryColumn[] = outputs.map((o) => {
        const spec = specs.get(o.colId);
        const summarize = o.role === "value" && !!spec && NUMBER_KINDS.has(spec.attribute.kind) && o.mType !== null && o.mType !== "type text";
        return { name: o.name, mType: o.mType, ...(summarize ? { summarize } : {}) };
    });
    if (model.recordUrlKey) columns.push({ name: recordUrlColumnName(settings.naming), mType: "type text", webUrl: true });
    return columns;
}

export interface DataModelInput {
    model: QueryModel;
    /** The main query's output columns. */
    outputs: OutputColumn[];
    mainName: string;
    /** One per choice column of the query (the choice tables already generated). */
    choiceQueries: ChoiceQuery[];
    /** The input the main model was built from, to build the bridges' lighter requests. */
    modelInput: ModelInput;
    settings: BuilderSettings;
    connectionUrl: string;
    /** Whether a column that isn't in the query can be added: in table mode and to FetchXML queries, not to aggregate ones. */
    canAddColumns: boolean;
    /** The primary key column's display name, for messages when it isn't in the query. */
    primaryKeyLabel?: string;
    /** Columns added to lookup tables besides key and name, by target table. */
    lookupColumns?: Record<string, LookupColumns>;
}

/** Lookups with more targets than this (Regarding, for example) get no lookup table. */
export const MAX_LOOKUP_TARGETS = 3;

/** Names for the usual combined lookup tables, by their sorted targets. */
const COMBINED_NAMES: Record<string, string> = { "systemuser,team": "Owner", "account,contact": "Customer" };

/**
 * The main query plus what it relates to: choice tables, bridge tables for multi-select columns, lookup
 * tables, and the relationships between them. Only columns in the main query's output take part.
 */
export function buildDataModel(input: DataModelInput): DataModel {
    const { model, outputs, settings, connectionUrl } = input;
    const mainName = cleanQueryName(input.mainName) || suggestedMainName(model.table);
    const inOutput = new Set(outputs.map((o) => o.colId));
    const choiceByAttribute = new Map(input.choiceQueries.map((q) => [`${q.entity}.${q.attribute}`, q]));

    const queries: ModelQuery[] = [];
    const relationships: Relationship[] = [];
    const notes: ModelNote[] = [];
    const usedPairs = new Set<string>();
    const relate = (from: Relationship["from"], to: Relationship["to"], bothDirections = false) => {
        // Power BI allows one active relationship between two tables; the others are kept for USERELATIONSHIP
        const pair = `${from.table}\u0000${to.table}`;
        const active = !usedPairs.has(pair);
        usedPairs.add(pair);
        relationships.push({ id: stableId(`${from.table}.${from.column}>${to.table}.${to.column}`), from, to, bothDirections, active });
    };

    // The main query's primary key column, which bridges relate to
    const pkSpec = model.columns.find((c) => c.attribute.isPrimaryId && !c.linkAlias && c.attribute.logicalName === model.table.primaryIdAttribute);
    const pkOutput = pkSpec ? outputs.find((o) => o.colId === pkSpec.id && o.sourceKey === pkSpec.key) : undefined;
    let booleans = 0;
    // Query names are unique in Power BI regardless of case
    const taken = new Set([mainName.toLowerCase(), ...input.choiceQueries.map((q) => q.queryName.toLowerCase())]);

    for (const q of input.choiceQueries) {
        queries.push({ id: q.id, kind: "choice", queryName: q.queryName, source: `${q.entity}.${q.attribute}`, code: q.code, columns: q.columns });
    }

    for (const spec of model.columns) {
        if (!inOutput.has(spec.id) || !isChoiceLike(spec.attribute.kind) || spec.aggregate) continue;
        const choice = choiceByAttribute.get(`${spec.entity}.${spec.attribute.logicalName}`);
        if (!choice) continue;
        const label = spec.attribute.displayName || spec.attribute.logicalName;

        if (spec.attribute.kind === "boolean") {
            booleans++;
            continue;
        }

        if (spec.attribute.kind === "multichoice") {
            if (spec.linkAlias) {
                notes.push({ text: `${label} is a multi-select column of a linked table. Bridge tables are only made for the main table's own columns.` });
                continue;
            }
            const bridge = bridgeModel(input, spec);
            if (!bridge) {
                notes.push({
                    text: `${label}: aggregate and distinct queries don't return the primary key, so there's no bridge table for this multi-select column.`,
                });
                continue;
            }
            const queryName = unique(taken, [bridgeQueryName(choice.queryName)]);
            const cols = bridgeColumns(queryName);
            queries.push({
                id: `bridge:${spec.id}`,
                kind: "bridge",
                queryName,
                source: `${spec.entity}.${spec.attribute.logicalName}`,
                code: generateBridgeQuery(bridge, spec, queryName, choice.queryName, settings, connectionUrl),
                columns: [
                    { name: cols.key, mType: "type text" },
                    { name: cols.value, mType: "Int64.Type" },
                ],
            });
            relate({ table: queryName, column: cols.value }, { table: choice.queryName, column: "Value" });
            if (pkOutput) {
                // Both directions, so a choice picked in the choice table filters the main table through the bridge
                relate({ table: queryName, column: cols.key }, { table: mainName, column: pkOutput.name }, true);
            } else {
                const canAdd = !!pkSpec || input.canAddColumns;
                notes.push({
                    text: `Add the primary key (${pkSpec?.attribute.displayName ?? input.primaryKeyLabel ?? model.table.primaryIdAttribute}) to the main query to relate ${queryName} to it.${canAdd ? "" : " Add it to the FetchXML first."}`,
                    fix: canAdd ? { kind: "include", colId: pkSpec?.id ?? `|${model.table.primaryIdAttribute}`, label: "Add the primary key" } : undefined,
                });
            }
            continue;
        }

        const valueColumn = outputs.find((o) => o.colId === spec.id && o.role === "value");
        if (valueColumn) {
            relate({ table: mainName, column: valueColumn.name }, { table: choice.queryName, column: "Value" });
        } else {
            notes.push({
                text: `${label} only returns its label, so it can't be related to ${choice.queryName}.`,
                fix: { kind: "mode", colId: spec.id, mode: "both", label: "Return label + value" },
            });
        }
    }
    if (booleans > 0) {
        notes.push({ text: `Yes/no columns aren't related to their choice tables: the label in the main query is all a report needs.` });
    }

    // Bridges directly after their choice table
    queries.sort((a, b) => a.source.localeCompare(b.source) || (a.kind === b.kind ? 0 : a.kind === "choice" ? -1 : 1));
    if (settings.lookupTables) addLookupTables(input, mainName, taken, queries, relate, notes);
    notes.push(...inactiveNotes(relationships));
    return { mainName, mainColumns: mainQueryColumns(input), queries, relationships, notes, tmdl: tmdlScript(relationships) };
}

/** The first name that isn't taken yet, else the first with a number; marks it taken. */
function unique(taken: Set<string>, candidates: string[]): string {
    const names = candidates.map(cleanQueryName).filter(Boolean);
    let name = names.find((n) => !taken.has(n.toLowerCase()));
    for (let i = 2; !name; i++) {
        if (!taken.has(`${names[0]} ${i}`.toLowerCase())) name = `${names[0]} ${i}`;
    }
    taken.add(name.toLowerCase());
    return name;
}

function andList(items: string[]): string {
    return items.length > 1 ? `${items.slice(0, -1).join(", ")} and ${items[items.length - 1]}` : items[0];
}

/** A target table's name in lookup table names and in the type column of a combined one. */
function tableLabel(t: TableInfo): string {
    return cleanQueryName(t.displayName) || t.logicalName;
}

/**
 * One lookup table per set of target tables, shared by every lookup column that points to the same
 * tables (Created By and Modified By both use User). The first relationship to it is active, the rest
 * inactive; the main table's Owner wins when it's one of them.
 */
function addLookupTables(
    input: DataModelInput,
    mainName: string,
    taken: Set<string>,
    queries: ModelQuery[],
    relate: (from: Relationship["from"], to: Relationship["to"]) => void,
    notes: ModelNote[],
): void {
    const { model, outputs } = input;
    const inOutput = new Set(outputs.map((o) => o.colId));
    const groups = new Map<string, { targets: TableInfo[]; uses: { spec: ColumnSpec; label: string; idColumn: string }[] }>();

    for (const spec of model.columns) {
        if (!inOutput.has(spec.id) || spec.attribute.kind !== "lookup" || spec.aggregate) continue;
        const label = spec.attribute.displayName || spec.attribute.logicalName;
        const targets = [...new Set(spec.attribute.targets ?? [])].sort();
        if (targets.length === 0) {
            notes.push({ text: `${label}: the tables it points to aren't known, so it has no lookup table.` });
            continue;
        }
        if (targets.length > MAX_LOOKUP_TARGETS) {
            notes.push({ text: `${label} can point to ${targets.length} tables, so it has no lookup table. Its name in the main query is what a report can use.` });
            continue;
        }
        const missing = targets.filter((t) => !input.modelInput.tables.has(t));
        if (missing.length > 0) {
            notes.push({ text: `${label} points to ${missing.join(", ")}, which isn't in the environment's table list, so it has no lookup table.` });
            continue;
        }
        const idColumn = outputs.find((o) => o.colId === spec.id && o.role === "id");
        if (!idColumn) {
            notes.push({
                text: `${label} only returns its name, so it can't be related to a lookup table.`,
                fix: { kind: "mode", colId: spec.id, mode: "both", label: "Return name + ID" },
            });
            continue;
        }
        const key = targets.join(",");
        const group = groups.get(key) ?? { targets: targets.map((t) => input.modelInput.tables.get(t)!), uses: [] };
        group.uses.push({ spec, label, idColumn: idColumn.name });
        groups.set(key, group);
    }

    const order = (u: { spec: ColumnSpec }) => (u.spec.linkAlias ? 2 : 0) + (u.spec.attribute.logicalName === "ownerid" ? 0 : 1);
    for (const [key, { targets, uses }] of groups) {
        uses.sort((a, b) => order(a) - order(b));
        const labels = uses.map((u) => u.label);
        let queryName: string;
        if (targets.length > 1) {
            const base = COMBINED_NAMES[key] ?? targets.map(tableLabel).join(" or ");
            queryName = unique(taken, [base, `${base} lookup`]);
        } else {
            const base = tableLabel(targets[0]);
            // A lookup to the main table itself (Parent Account) needs a table of its own: Power BI doesn't relate a table to itself
            const self = targets[0].logicalName === model.table.logicalName;
            const byColumn = uses.length === 1 ? [labels[0]] : [];
            queryName = unique(taken, self ? [...byColumn, `${base} lookup`] : [base, ...byColumn, `${base} lookup`]);
        }
        const cols = lookupColumns(queryName);
        const choice = targets.length === 1 ? input.lookupColumns?.[targets[0].logicalName] : undefined;
        const built = choice ? lookupTableQuery(targets[0], queryName, choice, labels, input) : null;
        const extended = built && built.added > 0 ? built : null;
        queries.push({
            id: `lookup:${key}`,
            kind: "lookup",
            queryName,
            source: `${targets.map((t) => t.logicalName).join(" + ")} for ${labels.join(", ")}`,
            code: generateLookupQuery(
                queryName,
                targets.map((t) => ({ table: t, label: tableLabel(t) })),
                labels,
                input.settings,
                input.connectionUrl,
            ),
            columns: [
                { name: cols.key, mType: "type text", hidden: true },
                { name: cols.name, mType: "type text" },
                ...(targets.length > 1 ? [{ name: cols.type, mType: "type text" }] : []),
            ],
            ...(targets.length === 1 ? { target: targets[0] } : {}),
            ...(extended ? { code: extended.code, columns: extended.columns, addedColumns: extended.added } : {}),
        });
        for (const u of uses) relate({ table: mainName, column: u.idColumn }, { table: queryName, column: cols.key });
    }
}

/** One note per table that's related to more than once, naming the columns whose relationship is inactive. */
function inactiveNotes(relationships: Relationship[]): ModelNote[] {
    const byTable = new Map<string, string[]>();
    for (const r of relationships) {
        if (!r.active) byTable.set(r.to.table, [...(byTable.get(r.to.table) ?? []), r.from.column]);
    }
    return [...byTable].map(([table, columns]) => ({
        text: `${andList(columns)} ${columns.length > 1 ? "relate" : "relates"} to ${table} through ${columns.length > 1 ? "inactive relationships" : "an inactive relationship"}: Power BI allows one active relationship between two tables. Use USERELATIONSHIP in a measure to filter by ${columns.length > 1 ? "them" : "it"}.`,
    }));
}

/**
 * The bridge's request: the main query's FetchXML with only the multi-select column left in, plus the
 * primary key. Built like the main query, so filters and joins stay the same. Null when the query can't
 * return the primary key (aggregate or distinct).
 */
function bridgeModel(input: DataModelInput, spec: ColumnSpec): QueryModel | null {
    const excluded = new Set(input.model.columns.filter((c) => c.id !== spec.id).map((c) => c.id));
    const m = buildQueryModel({ ...input.modelInput, includeRecordUrl: true, excluded, fromTablePicker: false, odataFilter: null });
    return m.recordUrlKey ? m : null;
}

export interface LookupTableQuery {
    model: QueryModel;
    outputs: OutputColumn[];
    /** The overrides the outputs were resolved with: the choice's, plus the fixed names of key and name. */
    overrides: ColumnOverrides;
    code: string;
    columns: QueryColumn[];
    /** Columns added besides key and name. */
    added: number;
}

/**
 * A lookup table of one table with more columns than key and name: the target table read through OData
 * like a table-picker query without a filter, so every row comes, with key and name under the lookup
 * table's own names and the added columns shaped like the main query's (labels, names, dates).
 * Null when the target table's columns aren't loaded yet.
 */
export function lookupTableQuery(
    target: TableInfo,
    queryName: string,
    choice: LookupColumns,
    usedBy: string[],
    input: Pick<DataModelInput, "modelInput" | "settings" | "connectionUrl">,
): LookupTableQuery | null {
    const attributes = input.modelInput.attributes.get(target.logicalName);
    if (!attributes) return null;
    const key = target.primaryIdAttribute;
    const name = target.primaryNameAttribute;
    const fixed = new Set([key, ...(name ? [name] : [])]);
    const extra = choice.columns.filter((c) => !fixed.has(c) && attributes.some((a) => a.logicalName === c));
    const model = buildQueryModel({
        doc: parseFetchXml(buildTableFetchXml(target, [...fixed, ...extra], false)),
        tables: input.modelInput.tables,
        attributes: input.modelInput.attributes,
        pageSize: input.modelInput.pageSize,
        includeRecordUrl: false,
        fromTablePicker: true,
        odataFilter: null,
    });
    const cols = lookupColumns(queryName);
    const overrides: ColumnOverrides = { ...choice.overrides, [`|${key}`]: { name: cols.key }, ...(name ? { [`|${name}`]: { name: cols.name } } : {}) };
    const outputs = resolveOutputColumns(model.columns, input.settings, overrides, false);
    const code = [
        `// Suggested query name: ${queryName}`,
        generateODataQuery(model, outputs, input.settings, input.connectionUrl, {
            headerLines: [`Lookup table: ${target.displayName} (${target.logicalName}) for ${usedBy.join(", ")}. Every row is read, active or not, so each record finds its match.`],
        }),
    ].join("\n");
    const columns = queryColumns(model, outputs, input.settings).map((c) => (c.name === cols.key ? { ...c, hidden: true } : c));
    return { model, outputs, overrides, code, columns, added: extra.length };
}
