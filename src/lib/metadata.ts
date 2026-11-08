import type { AttributeInfo, AttributeKind, DateBehavior, TableInfo } from "./types";

type Raw = Record<string, unknown>;

interface RawLabel {
    UserLocalizedLabel?: { Label?: string } | null;
    LocalizedLabels?: Array<{ Label?: string }>;
}

export function labelText(label: unknown): string | null {
    const l = label as RawLabel | null | undefined;
    if (!l) return null;
    const user = l.UserLocalizedLabel?.Label;
    if (user) return user;
    const first = l.LocalizedLabels?.find((x) => x.Label)?.Label;
    return first ?? null;
}

const KIND_BY_TYPENAME: Record<string, AttributeKind> = {
    StringType: "string",
    MemoType: "memo",
    IntegerType: "integer",
    BigIntType: "bigint",
    DecimalType: "decimal",
    DoubleType: "double",
    MoneyType: "money",
    BooleanType: "boolean",
    DateTimeType: "datetime",
    PicklistType: "choice",
    StateType: "state",
    StatusType: "status",
    MultiSelectPicklistType: "multichoice",
    LookupType: "lookup",
    CustomerType: "lookup",
    OwnerType: "lookup",
    UniqueidentifierType: "uniqueidentifier",
    EntityNameType: "entityname",
};

/** Columns that can't be returned in a query or make no sense in a report. */
const UNSUPPORTED_TYPENAMES = new Set(["ImageType", "FileType", "VirtualType", "PartyListType", "CalendarRulesType", "ManagedPropertyType"]);

export const ENTITY_PROPERTIES = [
    "LogicalName",
    "DisplayName",
    "EntitySetName",
    "PrimaryIdAttribute",
    "PrimaryNameAttribute",
    "IsCustomEntity",
    "IsIntersect",
    "IsActivity",
    "IsPrivate",
];

export const ATTRIBUTE_PROPERTIES = [
    "LogicalName",
    "SchemaName",
    "DisplayName",
    "AttributeType",
    "AttributeTypeName",
    "AttributeOf",
    "IsValidForRead",
    "IsPrimaryId",
    "IsPrimaryName",
    "IsCustomAttribute",
];

export function mapTables(raw: Raw[]): TableInfo[] {
    return raw
        .filter((e) => typeof e.EntitySetName === "string" && e.EntitySetName && e.IsPrivate !== true)
        .map((e) => ({
            logicalName: String(e.LogicalName),
            displayName: labelText(e.DisplayName) ?? String(e.LogicalName),
            entitySetName: String(e.EntitySetName),
            primaryIdAttribute: String(e.PrimaryIdAttribute ?? `${e.LogicalName}id`),
            primaryNameAttribute: (e.PrimaryNameAttribute as string | null) ?? null,
            isCustom: e.IsCustomEntity === true,
            isIntersect: e.IsIntersect === true,
            isActivity: e.IsActivity === true,
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName) || a.logicalName.localeCompare(b.logicalName));
}

function typeNameOf(a: Raw): string {
    const tn = a.AttributeTypeName as { Value?: string } | undefined;
    if (tn?.Value) return tn.Value;
    return `${String(a.AttributeType ?? "Unknown")}Type`;
}

export interface AttributeExtras {
    dateTimes?: Raw[];
    lookups?: Raw[];
}

export function mapAttributes(raw: Raw[], extras: AttributeExtras = {}): AttributeInfo[] {
    const dates = new Map<string, Raw>();
    for (const d of extras.dateTimes ?? []) dates.set(String(d.LogicalName), d);
    const lookups = new Map<string, Raw>();
    for (const l of extras.lookups ?? []) lookups.set(String(l.LogicalName), l);

    const result: AttributeInfo[] = [];
    for (const a of raw) {
        if (a.AttributeOf) continue;
        if (a.IsValidForRead === false) continue;
        const typeName = typeNameOf(a);
        if (UNSUPPORTED_TYPENAMES.has(typeName)) continue;
        const kind = KIND_BY_TYPENAME[typeName];
        if (!kind) continue;
        const logicalName = String(a.LogicalName);
        const info: AttributeInfo = {
            logicalName,
            schemaName: String(a.SchemaName ?? logicalName),
            displayName: labelText(a.DisplayName) ?? logicalName,
            typeName,
            kind,
            isPrimaryId: a.IsPrimaryId === true,
            isPrimaryName: a.IsPrimaryName === true,
            isCustom: a.IsCustomAttribute === true,
        };
        if (kind === "datetime") {
            const d = dates.get(logicalName);
            const behavior = (d?.DateTimeBehavior as { Value?: string } | undefined)?.Value;
            if (behavior === "UserLocal" || behavior === "DateOnly" || behavior === "TimeZoneIndependent") {
                info.dateBehavior = behavior as DateBehavior;
            }
            const format = d?.Format;
            if (format === "DateOnly" || format === "DateAndTime") info.dateFormat = format;
        }
        if (kind === "lookup") {
            const targets = lookups.get(logicalName)?.Targets;
            if (Array.isArray(targets)) info.targets = targets.map(String);
        }
        result.push(info);
    }
    return result.sort((x, y) => {
        // Primary id first, primary name second, then by display name
        const rank = (i: AttributeInfo) => (i.isPrimaryId ? 0 : i.isPrimaryName ? 1 : 2);
        return rank(x) - rank(y) || x.displayName.localeCompare(y.displayName) || x.logicalName.localeCompare(y.logicalName);
    });
}

export function kindLabel(a: AttributeInfo): string {
    switch (a.kind) {
        case "string":
            return "Text";
        case "memo":
            return "Multiline text";
        case "integer":
            return "Whole number";
        case "bigint":
            return "Big integer";
        case "decimal":
            return "Decimal";
        case "double":
            return "Float";
        case "money":
            return "Currency";
        case "boolean":
            return "Yes/No";
        case "datetime":
            if (a.dateBehavior === "DateOnly") return "Date only";
            if (a.dateBehavior === "TimeZoneIndependent") return a.dateFormat === "DateOnly" ? "Date (TZ independent)" : "Date/time (TZ independent)";
            return a.dateFormat === "DateOnly" ? "Date (user local)" : "Date and time";
        case "choice":
            return "Choice";
        case "state":
            return "Status";
        case "status":
            return "Status reason";
        case "multichoice":
            return "Choices";
        case "lookup": {
            const base = a.typeName === "CustomerType" ? "Customer" : a.typeName === "OwnerType" ? "Owner" : "Lookup";
            return a.targets && a.targets.length > 0 ? `${base} → ${a.targets.join(", ")}` : base;
        }
        case "uniqueidentifier":
            return a.isPrimaryId ? "Primary key" : "Unique identifier";
        case "entityname":
            return "Table name";
        default:
            return a.typeName.replace(/Type$/, "");
    }
}

export function isChoiceLike(kind: AttributeKind): boolean {
    return kind === "choice" || kind === "state" || kind === "status" || kind === "multichoice" || kind === "boolean" || kind === "entityname";
}

/** Metadata cast used by the Web API for each choice-like column type. */
export function optionMetadataType(kind: AttributeKind): string | null {
    switch (kind) {
        case "choice":
            return "PicklistAttributeMetadata";
        case "state":
            return "StateAttributeMetadata";
        case "status":
            return "StatusAttributeMetadata";
        case "multichoice":
            return "MultiSelectPicklistAttributeMetadata";
        case "boolean":
            return "BooleanAttributeMetadata";
        default:
            return null;
    }
}

export interface ChoiceOption {
    value: number | boolean;
    label: string;
    state?: number;
}

/** Turn an expanded *AttributeMetadata record into a list of options. */
export function mapOptions(kind: AttributeKind, raw: Raw): ChoiceOption[] {
    const os = ((raw.OptionSet as Raw | null) ?? (raw.GlobalOptionSet as Raw | null)) as Raw | null;
    if (!os) return [];
    if (kind === "boolean") {
        const t = os.TrueOption as Raw | undefined;
        const f = os.FalseOption as Raw | undefined;
        return [
            { value: true, label: labelText(t?.Label) ?? "Yes" },
            { value: false, label: labelText(f?.Label) ?? "No" },
        ];
    }
    const options = (os.Options as Raw[] | undefined) ?? [];
    return options.map((o) => {
        const opt: ChoiceOption = { value: Number(o.Value), label: labelText(o.Label) ?? String(o.Value) };
        if (kind === "status" && typeof o.State === "number") opt.state = o.State;
        return opt;
    });
}
