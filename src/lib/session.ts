import type { ColumnOverride, ColumnOverrides } from "./types";

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
    /** Column names and output modes, per table or FetchXML root table. */
    overridesByCtx: Record<string, ColumnOverrides>;
}

export const EMPTY_SESSION: Session = { mode: "table", selectedTable: null, selection: {}, fetchText: "", viewName: null, viewEdited: false, overridesByCtx: {} };

const VERSION = 1;
const MAX_FETCH_LENGTH = 200_000;

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
    return Object.keys(o).length > 0 ? o : null;
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
        overridesByCtx,
    };
}
