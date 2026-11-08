import { recordUrlColumnName } from "./columns";
import { normaliseUrl } from "./mcode";
import type { BuilderSettings, OutputColumn, QueryModel } from "./types";

export interface PreviewResult {
    columns: string[];
    rows: Array<Array<string | null>>;
    /** Keys in the response that the M query doesn't use (excluding annotations). */
    unusedKeys: string[];
}

function pad(n: number): string {
    return String(n).padStart(2, "0");
}

function formatDateTime(d: Date): string {
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}

/** Mirrors the M conversions so the preview shows what Power Query will produce. */
export function convertValue(col: OutputColumn, raw: unknown, settings: BuilderSettings): string | null {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === "string" && col.conversion) {
        if (col.conversion === "date") return raw.slice(0, 10);
        const d = new Date(raw);
        if (Number.isNaN(d.getTime())) return raw;
        if (col.conversion === "offset") d.setTime(d.getTime() + settings.utcOffsetMinutes * 60000);
        return formatDateTime(d);
    }
    if (typeof raw === "boolean") return raw ? "TRUE" : "FALSE";
    if (typeof raw === "object") return JSON.stringify(raw);
    return String(raw);
}

export function buildPreview(model: QueryModel, outputs: OutputColumn[], records: Record<string, unknown>[], settings: BuilderSettings, connectionUrl: string): PreviewResult {
    const columns = outputs.map((o) => o.name);
    const urlName = model.recordUrlKey ? recordUrlColumnName(settings.naming) : null;
    if (urlName) columns.push(urlName);
    const base = `${normaliseUrl(connectionUrl)}/main.aspx?${settings.appId.trim() ? `appid=${settings.appId.trim()}&` : ""}pagetype=entityrecord&etn=${model.table.logicalName}&id=`;

    const rows = records.map((r) => {
        const row = outputs.map((o) => convertValue(o, r[o.sourceKey], settings));
        if (urlName) {
            const id = r[model.recordUrlKey!];
            row.push(id ? base + String(id) : null);
        }
        return row;
    });

    const used = new Set(outputs.map((o) => o.sourceKey.split("@")[0]));
    if (model.recordUrlKey) used.add(model.recordUrlKey);
    const unused = new Set<string>();
    for (const r of records) {
        for (const k of Object.keys(r)) {
            if (k.includes("@")) continue;
            if (!used.has(k)) unused.add(k);
        }
    }
    // The Web API always adds the primary key (and currency for money columns); not worth flagging.
    unused.delete(model.table.primaryIdAttribute);
    unused.delete("_transactioncurrencyid_value");
    return { columns, rows, unusedKeys: Array.from(unused).sort() };
}
