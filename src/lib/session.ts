import type { ColumnOverride, ColumnOverrides, LookupColumns } from "./types";

/** What the builder remembers per environment, so reopening the tool continues where you stopped. */
export interface Session {
    mode: "table" | "fetchxml";
    selectedTable: string | null;
    /** Chosen columns per table (logical names). */
    selection: Record<string, string[]>;
    fetchText: string;
    /** The view the FetchXML was loaded from, if any, and whether it has been changed since. */
    viewName: string | null;
    viewEdited: boolean;
    /** The main query's name in the model, per table or FetchXML root table, when it differs from the suggestion. */
    queryNames: Record<string, string>;
    /** Column names and output modes, per table or FetchXML root table. */
    overridesByCtx: Record<string, ColumnOverrides>;
    /** Extra columns of lookup tables, per table or FetchXML root table, then per target table. */
    lookupColumnsByCtx: Record<string, Record<string, LookupColumns>>;
}

export const EMPTY_SESSION: Session = {
    mode: "table",
    selectedTable: null,
    selection: {},
    fetchText: "",
    viewName: null,
    viewEdited: false,
    queryNames: {},
    overridesByCtx: {},
    lookupColumnsByCtx: {},
};

const VERSION = 1;
/** Longer FetchXML isn't stored with the session; the builder says so. */
export const MAX_FETCH_LENGTH = 200_000;

export function sessionKey(environmentUrl: string): string {
    return `session:${environmentUrl.toLowerCase()}`;
}

export function serializeSession(session: Session): unknown {
    return { v: VERSION, ...session, fetchText: session.fetchText.length > MAX_FETCH_LENGTH ? "" : session.fetchText };
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseOverride(raw: unknown): ColumnOverride | null {
    if (!isRecord(raw)) return null;
    const o: ColumnOverride = {};
    if (typeof raw.mode === "string") o.mode = raw.mode;
    if (typeof raw.name === "string" && raw.name) o.name = raw.name;
    if (raw.hidden === true) o.hidden = true;
    if (raw.added === true) o.added = true;
    return Object.keys(o).length > 0 ? o : null;
}

function parseOverrides(raw: unknown): ColumnOverrides {
    const parsed: ColumnOverrides = {};
    if (!isRecord(raw)) return parsed;
    for (const [id, override] of Object.entries(raw)) {
        const o = parseOverride(override);
        if (o) parsed[id] = o;
    }
    return parsed;
}

function parseLookupColumns(raw: unknown): Session["lookupColumnsByCtx"] {
    const out: Session["lookupColumnsByCtx"] = {};
    if (!isRecord(raw)) return out;
    for (const [ctx, byTarget] of Object.entries(raw)) {
        if (!isRecord(byTarget)) continue;
        const targets: Record<string, LookupColumns> = {};
        for (const [target, choice] of Object.entries(byTarget)) {
            if (!isRecord(choice) || !Array.isArray(choice.columns)) continue;
            const columns = choice.columns.filter((c): c is string => typeof c === "string" && c !== "");
            if (columns.length > 0) targets[target] = { columns, overrides: parseOverrides(choice.overrides) };
        }
        if (Object.keys(targets).length > 0) out[ctx] = targets;
    }
    return out;
}

/** Reads a stored session, dropping anything that isn't in the expected shape. Returns null when there is nothing usable. */
export function parseSession(raw: unknown): Session | null {
    if (!isRecord(raw) || raw.v !== VERSION) return null;
    const selection: Session["selection"] = {};
    if (isRecord(raw.selection)) {
        for (const [table, columns] of Object.entries(raw.selection)) {
            if (Array.isArray(columns)) selection[table] = columns.filter((c): c is string => typeof c === "string");
        }
    }
    const overridesByCtx: Session["overridesByCtx"] = {};
    if (isRecord(raw.overridesByCtx)) {
        for (const [ctx, overrides] of Object.entries(raw.overridesByCtx)) {
            if (!isRecord(overrides)) continue;
            const parsed: ColumnOverrides = {};
            for (const [id, override] of Object.entries(overrides)) {
                const o = parseOverride(override);
                if (o) parsed[id] = o;
            }
            if (Object.keys(parsed).length > 0) overridesByCtx[ctx] = parsed;
        }
    }
    return {
        mode: raw.mode === "fetchxml" ? "fetchxml" : "table",
        selectedTable: typeof raw.selectedTable === "string" && raw.selectedTable ? raw.selectedTable : null,
        selection,
        fetchText: typeof raw.fetchText === "string" ? raw.fetchText : "",
        viewName: typeof raw.viewName === "string" && raw.viewName ? raw.viewName : null,
        viewEdited: raw.viewEdited === true,
        queryNames: isRecord(raw.queryNames)
            ? Object.fromEntries(Object.entries(raw.queryNames).filter((e): e is [string, string] => typeof e[1] === "string" && e[1].trim() !== ""))
            : {},
        overridesByCtx,
        lookupColumnsByCtx: parseLookupColumns(raw.lookupColumnsByCtx),
    };
}
