import { isChoiceLike } from "./metadata";
import type { BuilderSettings, ChoiceMode, ColumnOverrides, ColumnSpec, Conversion, DateMode, LookupMode, NamingMode, OutputColumn } from "./types";

export const FORMATTED = "@OData.Community.Display.V1.FormattedValue";
export const LOOKUP_TABLE = "@Microsoft.Dynamics.CRM.lookuplogicalname";

export type ModeOption = { value: string; label: string };

export const LOOKUP_MODES: ModeOption[] = [
    { value: "both", label: "Name + ID" },
    { value: "name", label: "Name" },
    { value: "id", label: "ID" },
    { value: "all", label: "Name + ID + table" },
];

export const CHOICE_MODES: ModeOption[] = [
    { value: "label", label: "Label" },
    { value: "value", label: "Value" },
    { value: "both", label: "Label + value" },
];

export const DATE_MODES: ModeOption[] = [
    { value: "utc", label: "UTC" },
    { value: "offset", label: "Local time" },
    { value: "formatted", label: "Formatted text" },
];

/** Which modes the column grid offers for a column, if any. */
export function modeOptionsFor(spec: ColumnSpec): ModeOption[] | null {
    const kind = spec.attribute.kind;
    if (kind === "lookup") return LOOKUP_MODES;
    if (isChoiceLike(kind)) return CHOICE_MODES;
    if (kind === "datetime" && spec.attribute.dateBehavior !== "DateOnly" && spec.attribute.dateBehavior !== "TimeZoneIndependent") return DATE_MODES;
    return null;
}

export function defaultModeFor(spec: ColumnSpec, settings: BuilderSettings): string | null {
    const kind = spec.attribute.kind;
    if (kind === "lookup") return settings.lookupMode;
    if (kind === "boolean") return settings.booleanMode;
    if (isChoiceLike(kind)) return settings.choiceMode;
    if (kind === "datetime") return settings.dateMode;
    return null;
}

/** Name used for the column before any role suffix, honouring the naming setting. */
export function defaultBaseName(spec: ColumnSpec, naming: NamingMode): string {
    if (spec.attrAlias) return spec.attrAlias;
    const a = spec.attribute;
    if (naming === "display") {
        let base = a.displayName || a.logicalName;
        if (a.isPrimaryId && !spec.linkAlias) base = /\bid$/i.test(base) ? base : `${base} ID`;
        return spec.linkLabel ? `${spec.linkLabel} ${base}` : base;
    }
    const name = naming === "schema" ? a.schemaName : a.logicalName;
    return spec.linkAlias ? `${spec.linkAlias}.${name}` : name;
}

function suffix(naming: NamingMode, kind: "id" | "name" | "table" | "value" | "label"): string {
    if (naming === "display") {
        return { id: " ID", name: "", table: " Table", value: " Value", label: "" }[kind];
    }
    const s = { id: "", name: "name", table: "type", value: "", label: "name" }[kind];
    return naming === "schema" && s ? s[0].toUpperCase() + s.slice(1) : s;
}

function valueType(spec: ColumnSpec): string | null {
    switch (spec.attribute.kind) {
        case "string":
        case "memo":
        case "uniqueidentifier":
        case "entityname":
        case "multichoice":
        case "lookup":
            return "type text";
        case "integer":
        case "bigint":
        case "choice":
        case "state":
        case "status":
            return "Int64.Type";
        case "decimal":
        case "double":
            return "type number";
        case "money":
            return "Currency.Type";
        case "boolean":
            return "type logical";
        default:
            return null;
    }
}

function dateConversion(spec: ColumnSpec, mode: DateMode): Conversion | "formatted" {
    const a = spec.attribute;
    if (a.dateBehavior === "DateOnly") return "date";
    if (a.dateBehavior === "TimeZoneIndependent") return a.dateFormat === "DateOnly" ? "date" : "wallclock";
    if (mode === "formatted") return "formatted";
    if (mode === "offset") return isLocalDateOnly(spec) ? "localdate" : "offset";
    return "utc";
}

/** A "Date only" format column with user-local behaviour: stored as midnight in the user's time zone, converted to UTC. */
function isLocalDateOnly(spec: ColumnSpec): boolean {
    const a = spec.attribute;
    return a.kind === "datetime" && a.dateFormat === "DateOnly" && a.dateBehavior !== "DateOnly" && a.dateBehavior !== "TimeZoneIndependent" && !spec.aggregate && !spec.dategrouping;
}

/** Things worth knowing about the chosen output that don't depend on the query itself. */
export function outputWarnings(specs: ColumnSpec[], outputs: OutputColumn[]): string[] {
    const byId = new Map(specs.map((s) => [s.id, s]));
    const shifted = outputs.filter((o) => o.conversion === "utc" && byId.has(o.colId) && isLocalDateOnly(byId.get(o.colId)!)).map((o) => `"${o.name}"`);
    if (shifted.length === 0) return [];
    const subject = shifted.length === 1 ? `${shifted[0]} is a date-only column` : `${shifted.join(", ")} are date-only columns`;
    return [`${subject} stored in each user's local time. In UTC the date can be a day off (for a user at UTC+1, midnight is 23:00 UTC the day before). Set Date/time to Local time to get the date users see.`];
}

export function recordUrlColumnName(naming: NamingMode): string {
    return naming === "display" ? "Record URL" : naming === "schema" ? "RecordUrl" : "recordurl";
}

/**
 * Expand the query's columns into the columns of the final table, applying the user's
 * output modes and names, and making every name unique.
 */
export function resolveOutputColumns(specs: ColumnSpec[], settings: BuilderSettings, overrides: ColumnOverrides): OutputColumn[] {
    const out: OutputColumn[] = [];
    const naming = settings.naming;
    for (const spec of specs) {
        const o = overrides[spec.id] ?? {};
        if (o.hidden) continue;
        const base = o.name?.trim() || defaultBaseName(spec, naming);
        const options = modeOptionsFor(spec);
        const mode = (o.mode && options?.some((m) => m.value === o.mode) ? o.mode : defaultModeFor(spec, settings)) ?? null;
        const kind = spec.attribute.kind;
        const push = (sourceKey: string, name: string, role: OutputColumn["role"], mType: string | null, conversion: Conversion | null = null) =>
            out.push({ colId: spec.id, sourceKey, name, role, mType, conversion });

        if (kind === "lookup") {
            const m = mode as LookupMode;
            if (m === "name" || m === "both" || m === "all") push(spec.key + FORMATTED, base + suffix(naming, "name"), "name", "type text");
            if (m === "id" || m === "both" || m === "all") push(spec.key, base + suffix(naming, "id"), "id", "type text");
            if (m === "all") push(spec.key + LOOKUP_TABLE, base + suffix(naming, "table"), "table", "type text");
        } else if (isChoiceLike(kind)) {
            const m = mode as ChoiceMode;
            if (m === "label" || m === "both") push(spec.key + FORMATTED, base + (m === "both" || naming !== "display" ? suffix(naming, "label") : ""), "label", "type text");
            if (m === "value" || m === "both") push(spec.key, base + (m === "both" ? suffix(naming, "value") : ""), "value", valueType(spec));
        } else if (kind === "datetime") {
            const conv = dateConversion(spec, (mode as DateMode) ?? "utc");
            if (conv === "formatted") push(spec.key + FORMATTED, base, "label", "type text");
            else push(spec.key, base, "value", conv === "date" || conv === "localdate" ? "type date" : "type datetime", conv);
        } else {
            push(spec.key, base, "value", valueType(spec));
        }
    }

    // Make names unique (Power Query is case-sensitive, but Power BI's model isn't)
    const taken = new Set<string>();
    const reserved = settings.includeRecordUrl ? recordUrlColumnName(naming).toLowerCase() : null;
    if (reserved) taken.add(reserved);
    for (const col of out) {
        let name = col.name;
        let n = 2;
        while (taken.has(name.toLowerCase())) {
            name = naming === "display" ? `${col.name} (${n})` : `${col.name}_${n}`;
            n += 1;
        }
        taken.add(name.toLowerCase());
        col.name = name;
    }
    return out;
}

export function needsAnnotation(outputs: OutputColumn[], annotation: string): boolean {
    return outputs.some((o) => o.sourceKey.endsWith(annotation));
}
