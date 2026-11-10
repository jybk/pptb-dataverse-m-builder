/** Normalised view of a Dataverse table, as the builder needs it. */
export interface TableInfo {
    logicalName: string;
    displayName: string;
    entitySetName: string;
    primaryIdAttribute: string;
    primaryNameAttribute: string | null;
    isCustom: boolean;
    isIntersect: boolean;
    isActivity: boolean;
}

export type AttributeKind =
    | "string"
    | "memo"
    | "integer"
    | "bigint"
    | "decimal"
    | "double"
    | "money"
    | "boolean"
    | "datetime"
    | "choice"
    | "state"
    | "status"
    | "multichoice"
    | "lookup"
    | "uniqueidentifier"
    | "entityname"
    | "unknown";

export type DateBehavior = "UserLocal" | "DateOnly" | "TimeZoneIndependent";

/** Normalised view of a Dataverse column. */
export interface AttributeInfo {
    logicalName: string;
    schemaName: string;
    displayName: string;
    /** AttributeTypeName.Value from metadata, e.g. "PicklistType". */
    typeName: string;
    kind: AttributeKind;
    isPrimaryId: boolean;
    isPrimaryName: boolean;
    isCustom: boolean;
    /** Lookup targets (logical names), when known. */
    targets?: string[];
    dateBehavior?: DateBehavior;
    dateFormat?: "DateOnly" | "DateAndTime";
}

export type LookupMode = "both" | "name" | "id" | "all";
export type ChoiceMode = "label" | "value" | "both";
export type DateMode = "utc" | "offset" | "formatted";
/** Daylight saving rule applied on top of the standard UTC offset when dates are shown in local time. */
export type DaylightSaving = "none" | "eu" | "us";
export type NamingMode = "display" | "logical" | "schema";
export type QueryStyle = "fetchxml" | "odata";

export interface BuilderSettings {
    /** Name of the Power Query parameter that holds the environment URL. */
    parameterName: string;
    /** When false the URL is written into each query as a constant instead of a parameter. */
    useParameter: boolean;
    apiVersion: string;
    pageSize: number;
    queryStyle: QueryStyle;
    activeOnly: boolean;
    includeRecordUrl: boolean;
    appId: string;
    naming: NamingMode;
    lookupMode: LookupMode;
    choiceMode: ChoiceMode;
    booleanMode: ChoiceMode;
    dateMode: DateMode;
    /** Standard-time offset from UTC in minutes, used when dateMode is "offset" (local time). */
    utcOffsetMinutes: number;
    daylightSaving: DaylightSaving;
    choiceTables: "live" | "static";
}

export const DEFAULT_SETTINGS: BuilderSettings = {
    parameterName: "DataverseUrl",
    useParameter: true,
    apiVersion: "v9.2",
    pageSize: 5000,
    queryStyle: "fetchxml",
    activeOnly: true,
    includeRecordUrl: false,
    appId: "",
    naming: "display",
    lookupMode: "both",
    choiceMode: "label",
    booleanMode: "value",
    dateMode: "utc",
    utcOffsetMinutes: 60,
    daylightSaving: "none",
    choiceTables: "live",
};

/** Per-column choices the user makes in the column grid. */
export interface ColumnOverride {
    /** Output mode for lookups / choices / dates; undefined = use the global default. */
    mode?: string;
    /** Replaces the default base column name. */
    name?: string;
    /** In FetchXML mode: leave this column out of the M output. */
    hidden?: boolean;
}

export type ColumnOverrides = Record<string, ColumnOverride>;

/** One attribute requested by the query, resolved against metadata. */
export interface ColumnSpec {
    /** Stable id: "<link alias or empty>|<attribute alias or logical name>". */
    id: string;
    entity: string;
    linkAlias: string | null;
    /** Display label used to prefix linked columns. */
    linkLabel: string | null;
    attrAlias: string | null;
    aggregate: string | null;
    dategrouping: string | null;
    attribute: AttributeInfo;
    /** Key of the value in the Web API JSON response. */
    key: string;
}

export type OutputRole = "value" | "label" | "id" | "name" | "table";
/** "localdate" is a user-local date-only column converted to local time and cut down to its date. */
export type Conversion = "utc" | "offset" | "date" | "wallclock" | "localdate";

/** One column of the final Power Query table. */
export interface OutputColumn {
    colId: string;
    sourceKey: string;
    name: string;
    role: OutputRole;
    /** M type expression, e.g. "type text" or "Int64.Type". Null = leave untyped. */
    mType: string | null;
    conversion: Conversion | null;
}

/** Everything the M generator needs for one query. */
export interface QueryModel {
    table: TableInfo;
    /** Normalised FetchXML without paging attributes, one element per line (indent, text). */
    fetchLines: Array<{ depth: number; text: string }>;
    fetchXml: string;
    columns: ColumnSpec[];
    /** True when the FetchXML has `top` or is an aggregate query: a single request, no paging. */
    singleRequest: boolean;
    pageSize: number;
    isAggregate: boolean;
    hasLinks: boolean;
    /** Primary key JSON key, when a record URL can be built. */
    recordUrlKey: string | null;
    /** Only for queries built from the table picker: OData equivalent. */
    odata: { select: string[]; filter: string | null } | null;
    warnings: string[];
}
