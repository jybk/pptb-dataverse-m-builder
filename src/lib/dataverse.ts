import { ATTRIBUTE_PROPERTIES, ENTITY_PROPERTIES, mapAttributes, mapOptions, mapTables, optionMetadataType, type ChoiceOption } from "./metadata";
import type { AttributeInfo, AttributeKind, TableInfo } from "./types";

type Raw = Record<string, unknown>;

function api() {
    return window.dataverseAPI;
}

function entityPath(entity: string): string {
    return `EntityDefinitions(LogicalName='${encodeURIComponent(entity)}')`;
}

export function errorMessage(err: unknown): string {
    if (err instanceof Error) return err.message;
    if (typeof err === "string") return err;
    try {
        return JSON.stringify(err);
    } catch {
        return String(err);
    }
}

export async function loadTables(): Promise<TableInfo[]> {
    const res = await api().getAllEntitiesMetadata(ENTITY_PROPERTIES);
    return mapTables(res.value as Raw[]);
}

async function optional(query: string): Promise<Raw[]> {
    try {
        const res = await api().queryData(query);
        return (res.value ?? []) as Raw[];
    } catch (err) {
        console.warn("Optional metadata query failed", query, err);
        return [];
    }
}

export async function loadAttributes(entity: string): Promise<AttributeInfo[]> {
    const [attrs, dateTimes, lookups] = await Promise.all([
        api().getEntityRelatedMetadata(entity, "Attributes", ATTRIBUTE_PROPERTIES),
        optional(`${entityPath(entity)}/Attributes/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=LogicalName,Format,DateTimeBehavior`),
        optional(`${entityPath(entity)}/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,Targets`),
    ]);
    return mapAttributes(((attrs as { value?: Raw[] }).value ?? []) as Raw[], { dateTimes, lookups });
}

export interface ViewInfo {
    id: string;
    name: string;
    fetchXml: string;
    personal: boolean;
    isDefault: boolean;
}

export async function loadViews(entity: string): Promise<ViewInfo[]> {
    const filter = encodeURIComponent(`returnedtypecode eq '${entity}' and querytype eq 0`);
    // System views are the point of the menu, so a failure there is reported; personal views are a bonus
    const [system, personal] = await Promise.all([
        api()
            .queryData(`savedqueries?$select=savedqueryid,name,fetchxml,isdefault&$filter=${filter}&$orderby=name`)
            .then((res) => (res.value ?? []) as Raw[]),
        optional(`userqueries?$select=userqueryid,name,fetchxml&$filter=${filter}&$orderby=name`),
    ]);
    return [
        ...system.map((v) => ({ id: String(v.savedqueryid), name: String(v.name), fetchXml: String(v.fetchxml ?? ""), personal: false, isDefault: v.isdefault === true })),
        ...personal.map((v) => ({ id: String(v.userqueryid), name: String(v.name), fetchXml: String(v.fetchxml ?? ""), personal: true, isDefault: false })),
    ].filter((v) => v.fetchXml);
}

/** Options for every choice column of one kind on a table, keyed by column logical name. */
export async function loadOptions(entity: string, kind: AttributeKind): Promise<Map<string, ChoiceOption[]>> {
    const cast = optionMetadataType(kind);
    const result = new Map<string, ChoiceOption[]>();
    if (!cast) return result;
    const hasGlobal = kind === "choice" || kind === "multichoice";
    const expand = kind === "boolean" ? "OptionSet" : hasGlobal ? "OptionSet($select=Options),GlobalOptionSet($select=Options)" : "OptionSet($select=Options)";
    const res = await api().queryData(`${entityPath(entity)}/Attributes/Microsoft.Dynamics.CRM.${cast}?$select=LogicalName&$expand=${expand}`);
    for (const a of (res.value ?? []) as Raw[]) result.set(String(a.LogicalName), mapOptions(kind, a));
    return result;
}

/** Runs the query for the preview, using the real entity set name. */
export async function runFetchPreview(entitySetName: string, fetchXml: string): Promise<Record<string, unknown>[]> {
    const res = await api().queryData(`${entitySetName}?fetchXml=${encodeURIComponent(fetchXml)}`);
    return (res.value ?? []) as Record<string, unknown>[];
}

export async function runODataPreview(entitySetName: string, select: string[], filter: string | null, top: number): Promise<Record<string, unknown>[]> {
    let q = `${entitySetName}?$select=${select.map(encodeURIComponent).join(",")}&$top=${top}`;
    if (filter) q += `&$filter=${encodeURIComponent(filter)}`;
    const res = await api().queryData(q);
    return (res.value ?? []) as Record<string, unknown>[];
}

export interface ModelDrivenApp {
    id: string;
    name: string;
}

export async function loadApps(): Promise<ModelDrivenApp[]> {
    const rows = await optional(`appmodules?$select=appmoduleid,name&$orderby=name`);
    return rows.map((r) => ({ id: String(r.appmoduleid), name: String(r.name) }));
}
