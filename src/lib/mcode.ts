import { FORMATTED, LOOKUP_TABLE, needsAnnotation, recordUrlColumnName } from "./columns";
import type { ChoiceOption } from "./metadata";
import { optionMetadataType } from "./metadata";
import { offsetLabel } from "./timezone";
import type { AttributeKind, BuilderSettings, ColumnSpec, OutputColumn, QueryModel, TableInfo } from "./types";

const PAGING_COOKIE = "Microsoft.Dynamics.CRM.fetchxmlpagingcookie";
const MORE_RECORDS = "Microsoft.Dynamics.CRM.morerecords";
const M_KEYWORDS = new Set([
    "and",
    "as",
    "each",
    "else",
    "error",
    "false",
    "if",
    "in",
    "is",
    "let",
    "meta",
    "not",
    "null",
    "or",
    "otherwise",
    "section",
    "shared",
    "then",
    "true",
    "try",
    "type",
]);

/** M text literal: doubles quotes and escapes characters M would otherwise interpret. */
export function mText(value: string): string {
    const escaped = value.replace(/"/g, '""').replace(/#\(/g, "#(#)(").replace(/\r/g, "#(cr)").replace(/\n/g, "#(lf)").replace(/\t/g, "#(tab)");
    return `"${escaped}"`;
}

/** M identifier, quoted with #"..." when needed. */
export function mIdent(name: string): string {
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !M_KEYWORDS.has(name) ? name : `#${mText(name)}`;
}

/**
 * Environment URL for Power Query: no trailing slash or path, and the organisation host rather than
 * the Web API host ToolBox may hand us (org.api.crm4.dynamics.com → org.crm4.dynamics.com), because
 * record links only open on the organisation host and it's the usual sign-in resource.
 */
export function normaliseUrl(url: string): string {
    const trimmed = url.trim();
    try {
        const u = new URL(trimmed);
        const labels = u.hostname.split(".");
        if (labels.length >= 4 && labels[1].toLowerCase() === "api" && /^crm\d*$/i.test(labels[2])) labels.splice(1, 1);
        return `${u.protocol}//${labels.join(".")}${u.port ? `:${u.port}` : ""}`;
    } catch {
        return trimmed.replace(/\/+$/, "");
    }
}

const I1 = "    ";
const I2 = I1 + I1;
const I3 = I2 + I1;
const I4 = I3 + I1;

function list(items: string[], indent: string): string {
    if (items.length === 0) return "{}";
    return `{\n${items.map((i) => indent + I1 + i).join(",\n")}\n${indent}}`;
}

function preferHeader(annotations: string[], extra: string[] = []): string | null {
    const parts: string[] = [];
    if (annotations.length > 0) parts.push(`odata.include-annotations="${annotations.join(",")}"`);
    parts.push(...extra);
    return parts.length ? parts.join(",") : null;
}

/**
 * Request headers. With a credential (Organizational account) Power Query only allows Accept,
 * Accept-Charset, Accept-Encoding, Accept-Language, Cache-Control, Content-Type, If-Modified-Since,
 * Prefer, Range and Referer, so no OData-Version headers here (Dataverse defaults to OData 4.0).
 */
export const ALLOWED_AUTH_HEADERS = [
    "Accept",
    "Accept-Charset",
    "Accept-Encoding",
    "Accept-Language",
    "Cache-Control",
    "Content-Type",
    "If-Modified-Since",
    "Prefer",
    "Range",
    "Referer",
];

function headersRecord(prefer: string | null, indent: string): string {
    const fields = [`Accept = "application/json"`];
    if (prefer) fields.push(`Prefer = ${mText(prefer)}`);
    return `[\n${fields.map((f) => indent + I1 + f).join(",\n")}\n${indent}]`;
}

function urlRef(settings: BuilderSettings): string {
    return mIdent(settings.parameterName.trim() || "DataverseUrl");
}

function header(lines: string[], settings: BuilderSettings, connectionUrl: string): string[] {
    const out = lines.map((l) => `// ${l}`);
    if (settings.useParameter) {
        out.push(`// Needs a text parameter named ${settings.parameterName.trim() || "DataverseUrl"} = ${serviceRootUrl(connectionUrl, settings)} (no trailing slash)`);
    }
    return out;
}

function urlStep(settings: BuilderSettings, connectionUrl: string): string[] {
    if (settings.useParameter) return [];
    return [`${I1}${urlRef(settings)} = ${mText(serviceRootUrl(connectionUrl, settings))},`];
}

/**
 * Value of the URL parameter: the environment's Web API root, e.g. https://org.crm4.dynamics.com/api/data/v9.2.
 * Every query passes the parameter as-is as the first argument of Web.Contents, which keeps the API
 * version in one place, lets Power BI offer "Organizational account" (the bare environment URL doesn't
 * answer with an OAuth challenge), and keeps scheduled refresh working (no dynamic data source).
 */
export function serviceRootUrl(connectionUrl: string, settings: BuilderSettings): string {
    const env = normaliseUrl(connectionUrl) || "https://yourorg.crm.dynamics.com";
    return `${env}/api/data/${settings.apiVersion.trim() || "v9.2"}`;
}

function offsetParts(minutes: number): [number, number] {
    const h = Math.trunc(minutes / 60);
    const m = minutes % 60;
    return [h, m === 0 ? 0 : m];
}

function duration(minutes: number): string {
    const [h, m] = offsetParts(minutes);
    return `#duration(0, ${h}, ${m}, 0)`;
}

/**
 * The ToLocalDateTime function: UTC text from Dataverse to a datetime on the local clock.
 * The daylight saving rules are the ones in timezone.ts, which the preview uses.
 */
function localTimeSteps(settings: BuilderSettings): string[] {
    const standard = settings.utcOffsetMinutes;
    const [h, m] = offsetParts(standard);
    if (settings.daylightSaving === "eu") {
        return [
            `${I1}// Local time: ${offsetLabel(standard)}, and ${offsetLabel(standard + 60)} during European summer time`,
            `${I1}// (last Sunday of March to last Sunday of October; the clocks change at 01:00 UTC)`,
            `${I1}LastSundayOf = (year as number, month as number) as datetime =>`,
            `${I2}let LastDay = Date.EndOfMonth(#date(year, month, 1))`,
            `${I2}in DateTime.From(Date.AddDays(LastDay, -Date.DayOfWeek(LastDay, Day.Sunday))) + #duration(0, 1, 0, 0),`,
            `${I1}ToLocalDateTime = (value) =>`,
            `${I2}if value = null then null`,
            `${I2}else`,
            `${I3}let`,
            `${I4}Utc = DateTimeZone.RemoveZone(DateTimeZone.FromText(value)),`,
            `${I4}IsSummerTime = Utc >= LastSundayOf(Date.Year(Utc), 3) and Utc < LastSundayOf(Date.Year(Utc), 10)`,
            `${I3}in`,
            `${I4}Utc + (if IsSummerTime then ${duration(standard + 60)} else ${duration(standard)}),`,
        ];
    }
    if (settings.daylightSaving === "us") {
        return [
            `${I1}// Local time: ${offsetLabel(standard)}, and ${offsetLabel(standard + 60)} during daylight saving time (United States and Canada:`,
            `${I1}// second Sunday of March to first Sunday of November; the clocks change at 02:00 local time)`,
            `${I1}NthSundayOf = (year as number, month as number, n as number) as datetime =>`,
            `${I2}let First = #date(year, month, 1)`,
            `${I2}in DateTime.From(Date.AddDays(First, Number.Mod(7 - Date.DayOfWeek(First, Day.Sunday), 7) + 7 * (n - 1))),`,
            `${I1}ToLocalDateTime = (value) =>`,
            `${I2}if value = null then null`,
            `${I2}else`,
            `${I3}let`,
            `${I4}Standard = DateTimeZone.RemoveZone(DateTimeZone.FromText(value)) + ${duration(standard)},`,
            `${I4}Year = Date.Year(Standard),`,
            `${I4}IsDaylightSaving = Standard >= NthSundayOf(Year, 3, 2) + #duration(0, 2, 0, 0) and Standard < NthSundayOf(Year, 11, 1) + #duration(0, 1, 0, 0)`,
            `${I3}in`,
            `${I4}if IsDaylightSaving then Standard + #duration(0, 1, 0, 0) else Standard,`,
        ];
    }
    return [
        `${I1}// Local time: ${offsetLabel(standard)} all year. Choose a daylight saving rule in the builder if the clocks change where you are.`,
        `${I1}ToLocalDateTime = (value) => if value = null then null else DateTimeZone.RemoveZone(DateTimeZone.SwitchZone(DateTimeZone.FromText(value), ${h}, ${m})),`,
    ];
}

/**
 * Steps shared by both query styles: from the list of JSON records to the final typed table.
 * All column decisions live in one `Columns` list (response key, name, type, optional conversion),
 * so a column is added, removed or renamed by editing one line.
 */
function shapeSteps(model: QueryModel, outputs: OutputColumn[], settings: BuilderSettings): { steps: string[]; last: string } {
    const steps: string[] = [""];
    const conversions = new Set(outputs.map((o) => o.conversion).filter(Boolean));
    if (conversions.has("utc") || conversions.has("wallclock")) {
        steps.push(`${I1}ToDateTime = (value) => if value = null then null else DateTimeZone.RemoveZone(DateTimeZone.FromText(value)),`);
    }
    if (conversions.has("offset") || conversions.has("localdate")) steps.push(...localTimeSteps(settings));
    if (conversions.has("localdate")) {
        steps.push(`${I1}// Date-only columns stored in local time: convert to local time first, then keep the date`, `${I1}ToLocalDate = (value) => Date.From(ToLocalDateTime(value)),`);
    }
    if (conversions.has("date")) {
        steps.push(
            `${I1}ToDate = (value) => if value = null then null else #date(Number.From(Text.Range(value, 0, 4)), Number.From(Text.Range(value, 5, 2)), Number.From(Text.Range(value, 8, 2))),`,
        );
    }
    const usesFormatted = outputs.some((o) => o.sourceKey.endsWith(FORMATTED));
    const usesLookupTable = outputs.some((o) => o.sourceKey.endsWith(LOOKUP_TABLE));
    if (usesFormatted) steps.push(`${I1}Formatted = ${mText(FORMATTED)},`);
    if (usesLookupTable) steps.push(`${I1}LookupTable = ${mText(LOOKUP_TABLE)},`);

    const keyExpr = (key: string) =>
        key.endsWith(FORMATTED)
            ? `${mText(key.slice(0, -FORMATTED.length))} & Formatted`
            : key.endsWith(LOOKUP_TABLE)
              ? `${mText(key.slice(0, -LOOKUP_TABLE.length))} & LookupTable`
              : mText(key);
    const CONVERSION_FUNCTIONS = { utc: "ToDateTime", wallclock: "ToDateTime", offset: "ToLocalDateTime", localdate: "ToLocalDate", date: "ToDate" } as const;
    const convFn = (o: OutputColumn) => (o.conversion ? CONVERSION_FUNCTIONS[o.conversion] : null);
    const cells = outputs.map((o) => {
        const fn = convFn(o);
        return [keyExpr(o.sourceKey) + ",", mText(o.name) + ",", (o.mType ?? "type any") + (fn ? "," : ""), fn ?? ""];
    });
    const width = (i: number, cap: number) => Math.min(cap, Math.max(0, ...cells.map((c) => c[i].length)));
    const [w0, w1, w2] = [width(0, 64), width(1, 56), width(2, 16)];
    const rows = cells.map((c) => (c[3] ? `{${c[0].padEnd(w0)} ${c[1].padEnd(w1)} ${c[2].padEnd(w2)} ${c[3]}}` : `{${c[0].padEnd(w0)} ${c[1].padEnd(w1)} ${c[2]}}`.trimEnd()));
    steps.push(`${I1}// Columns: response key, column name, type and (for dates) conversion. Edit a line to rename or retype a column.`, `${I1}Columns = ${list(rows, I1)},`);

    const urlKey = model.recordUrlKey;
    const needsHelperKey = urlKey !== null && !outputs.some((o) => o.sourceKey === urlKey);
    steps.push(`${I1}Source = Table.FromRecords(Records, List.Transform(Columns, each _{0})${needsHelperKey ? ` & {${mText(urlKey!)}}` : ""}, MissingField.UseNull),`);
    let prev = "Source";

    const recordUrlName = recordUrlColumnName(settings.naming);
    if (urlKey) {
        const appPart = settings.appId.trim() ? `appid=${settings.appId.trim()}&` : "";
        steps.push(
            `${I1}// The record link uses the environment part of the URL parameter (everything before /api/data/)`,
            `${I1}EnvironmentUrl = let Position = Text.PositionOf(Text.Lower(${urlRef(settings)}), "/api/data/") in Text.TrimEnd(if Position < 0 then ${urlRef(settings)} else Text.Start(${urlRef(settings)}, Position), "/"),`,
            `${I1}RecordUrlBase = EnvironmentUrl & ${mText(`/main.aspx?${appPart}pagetype=entityrecord&etn=${model.table.logicalName}&id=`)},`,
            `${I1}WithRecordUrl = Table.AddColumn(${prev}, ${mText(recordUrlName)}, (row) =>`,
            `${I2}let Id = Record.Field(row, ${mText(urlKey)}) in if Id = null then null else RecordUrlBase & Id, type text),`,
        );
        prev = "WithRecordUrl";
    }

    const sameName = outputs.some((o) => o.sourceKey === o.name);
    const renamePairs = "List.Transform(Columns, each {_{0}, _{1}})";
    steps.push(`${I1}Renamed = Table.RenameColumns(${prev}, ${sameName ? `List.Select(${renamePairs}, each _{0} <> _{1})` : renamePairs}),`);
    prev = "Renamed";

    if (conversions.size > 0) {
        steps.push(`${I1}Converted = Table.TransformColumns(${prev}, List.Transform(List.Select(Columns, each List.Count(_) > 3), each {_{1}, _{3}})),`);
        prev = "Converted";
    }

    steps.push(`${I1}Typed = Table.TransformColumnTypes(${prev}, List.Transform(Columns, each {_{1}, _{2}})),`);
    prev = "Typed";

    if (needsHelperKey) {
        steps.push(`${I1}Result = Table.SelectColumns(${prev}, List.Transform(Columns, each _{1}) & {${mText(recordUrlName)}}),`);
        prev = "Result";
    }

    // Drop the trailing comma of the last step
    steps[steps.length - 1] = steps[steps.length - 1].replace(/,$/, "");
    return { steps, last: prev };
}

function fetchXmlStep(model: QueryModel): string[] {
    const lines = model.fetchLines.map((l) => I2 + I1.repeat(l.depth) + mText(l.text));
    return [`${I1}FetchXml = Text.Combine({`, lines.join(",\n"), `${I1}}),`];
}

function summary(model: QueryModel, outputs: OutputColumn[], style: string): string {
    const cols = `${outputs.length} column${outputs.length === 1 ? "" : "s"}`;
    return `Table: ${model.table.displayName} (${model.table.logicalName}) · ${cols} · ${style}`;
}

/** M query that sends the FetchXML through the Web API, following paging cookies to get every row. */
export function generateFetchXmlQuery(model: QueryModel, outputs: OutputColumn[], settings: BuilderSettings, connectionUrl: string): string {
    const annotations: string[] = [];
    if (needsAnnotation(outputs, FORMATTED)) annotations.push(FORMATTED.slice(1));
    if (needsAnnotation(outputs, LOOKUP_TABLE)) annotations.push(LOOKUP_TABLE.slice(1));
    const shape = shapeSteps(model, outputs, settings);
    const style = model.singleRequest ? "FetchXML, single request" : `FetchXML, ${model.pageSize} rows per request with paging cookies`;

    const out: string[] = [
        ...header(["Generated by Dataverse M Builder for Power Platform ToolBox", summary(model, outputs, style)], settings, connectionUrl),
        "let",
        ...urlStep(settings, connectionUrl),
        `${I1}EntitySetName = ${mText(model.table.entitySetName)},`,
    ];
    if (!model.singleRequest) out.push(`${I1}PageSize = ${model.pageSize},`);
    out.push(...fetchXmlStep(model));

    if (model.singleRequest) {
        const prefer = preferHeader(annotations);
        out.push(
            "",
            `${I1}Response = Json.Document(Web.Contents(${urlRef(settings)}, [`,
            `${I2}RelativePath = EntitySetName,`,
            `${I2}Query = [fetchXml = FetchXml],`,
            `${I2}Headers = ${headersRecord(prefer, I2)}`,
            `${I1}])),`,
            `${I1}Records = Response[value],`,
        );
    } else {
        const prefer = preferHeader([...annotations, PAGING_COOKIE, MORE_RECORDS]);
        out.push(
            "",
            `${I1}// Paging: the cookie comes back URL-encoded twice and goes into the next request XML-encoded`,
            `${I1}UrlDecode = (value as text) as text => Uri.Parts("http://localhost/?v=" & value)[Query][v],`,
            `${I1}XmlEncode = (value as text) as text =>`,
            `${I2}Text.Replace(Text.Replace(Text.Replace(Text.Replace(value, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;"),`,
            `${I1}GetPage = (page as number, pagingCookie as nullable text) as record =>`,
            `${I2}let`,
            `${I3}PagedFetchXml = "<fetch page=""" & Text.From(page) & """ count=""" & Text.From(PageSize) & """"`,
            `${I4}& (if pagingCookie = null then "" else " paging-cookie=""" & XmlEncode(pagingCookie) & """")`,
            `${I4}& Text.Middle(FetchXml, 6),`,
            `${I3}Response = Json.Document(Web.Contents(${urlRef(settings)}, [`,
            `${I4}RelativePath = EntitySetName,`,
            `${I4}Query = [fetchXml = PagedFetchXml],`,
            `${I4}Headers = ${headersRecord(prefer, I4)}`,
            `${I3}])),`,
            `${I3}CookieXml = Record.FieldOrDefault(Response, ${mText("@" + PAGING_COOKIE)}, null),`,
            `${I3}NextCookie = if CookieXml = null then null`,
            `${I4}else UrlDecode(UrlDecode(Text.BetweenDelimiters(CookieXml, "pagingcookie=""", """")))`,
            `${I2}in`,
            `${I3}[`,
            `${I4}Records = Response[value],`,
            `${I4}MoreRecords = Record.FieldOrDefault(Response, ${mText("@" + MORE_RECORDS)}, false) = true,`,
            `${I4}Page = page,`,
            `${I4}PagingCookie = NextCookie`,
            `${I3}],`,
            `${I1}Pages = List.Generate(`,
            `${I2}() => GetPage(1, null),`,
            `${I2}each _ <> null,`,
            `${I2}each if [MoreRecords] then GetPage([Page] + 1, [PagingCookie]) else null,`,
            `${I2}each [Records]`,
            `${I1}),`,
            `${I1}Records = List.Combine(Pages),`,
        );
    }
    out.push(...shape.steps, "in", `${I1}${shape.last}`);
    return out.join("\n");
}

/** Columns for the OData $select: the base key of every output column, plus the primary key for record URLs. */
export function odataSelect(model: QueryModel, outputs: OutputColumn[]): string[] {
    return Array.from(new Set(outputs.map((o) => o.sourceKey.split("@")[0]).concat(model.recordUrlKey ? [model.recordUrlKey] : [])));
}

/** M query using an OData $select (single table only), following @odata.nextLink. */
export function generateODataQuery(model: QueryModel, outputs: OutputColumn[], settings: BuilderSettings, connectionUrl: string): string {
    if (!model.odata) throw new Error("This query can't be expressed in OData.");
    const annotations: string[] = [];
    if (needsAnnotation(outputs, FORMATTED)) annotations.push(FORMATTED.slice(1));
    if (needsAnnotation(outputs, LOOKUP_TABLE)) annotations.push(LOOKUP_TABLE.slice(1));
    const prefer = preferHeader(annotations, [`odata.maxpagesize=${model.pageSize}`]);
    const shape = shapeSteps(model, outputs, settings);

    const select = odataSelect(model, outputs);
    const queryFields = [`#"$select" = ${mText(select.join(","))}`];
    if (model.odata.filter) queryFields.push(`#"$filter" = ${mText(model.odata.filter)}`);

    const out: string[] = [
        ...header(
            ["Generated by Dataverse M Builder for Power Platform ToolBox", summary(model, outputs, `OData, ${model.pageSize} rows per request`)],
            settings,
            connectionUrl,
        ),
        "let",
        ...urlStep(settings, connectionUrl),
        `${I1}EntitySetName = ${mText(model.table.entitySetName)},`,
        `${I1}FirstPageQuery = [`,
        queryFields.map((f) => I2 + f).join(",\n"),
        `${I1}],`,
        "",
        `${I1}GetPage = (query as record) as record =>`,
        `${I2}Json.Document(Web.Contents(${urlRef(settings)}, [`,
        `${I3}RelativePath = EntitySetName,`,
        `${I3}Query = query,`,
        `${I3}Headers = ${headersRecord(prefer, I3)}`,
        `${I2}])),`,
        `${I1}// Follow @odata.nextLink, re-using its query string so refresh in the Power BI service keeps working`,
        `${I1}Pages = List.Generate(`,
        `${I2}() => GetPage(FirstPageQuery),`,
        `${I2}each _ <> null,`,
        `${I2}each if Record.HasFields(_, "@odata.nextLink") then GetPage(Uri.Parts(Record.Field(_, "@odata.nextLink"))[Query]) else null,`,
        `${I2}each _[value]`,
        `${I1}),`,
        `${I1}Records = List.Combine(Pages),`,
        ...shape.steps,
        "in",
        `${I1}${shape.last}`,
    ];
    return out.join("\n");
}

/** Blank-query body for the environment URL parameter. */
export function generateParameterQuery(connectionUrl: string, settings: BuilderSettings): string {
    return `${mText(serviceRootUrl(connectionUrl, settings))} meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]`;
}

export interface ChoiceQuery {
    id: string;
    queryName: string;
    entity: string;
    attribute: string;
    kind: AttributeKind;
    code: string;
}

/** Columns that get a choice table (one per distinct table + column). */
export function choiceColumns(columns: ColumnSpec[]): ColumnSpec[] {
    const seen = new Set<string>();
    return columns.filter((c) => {
        if (!optionMetadataType(c.attribute.kind) || c.aggregate) return false;
        const k = `${c.entity}.${c.attribute.logicalName}`;
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
    });
}

export function choiceQueryName(table: TableInfo | undefined, spec: ColumnSpec): string {
    const t = table?.displayName ?? spec.entity;
    return `${t} ${spec.attribute.displayName}`
        .replace(/[^\p{L}\p{N} _-]/gu, "")
        .replace(/\s+/g, " ")
        .trim();
}

function staticChoice(spec: ColumnSpec, options: ChoiceOption[]): string[] {
    const kind = spec.attribute.kind;
    if (kind === "boolean") {
        return [
            `${I1}Source = #table(type table [Value = logical, Label = text], ${list(
                options.map((o) => `{${o.value ? "true" : "false"}, ${mText(o.label)}}`),
                I1,
            )})`,
        ];
    }
    if (kind === "status") {
        return [
            `${I1}Source = #table(type table [Value = Int64.Type, Label = text, State = Int64.Type], ${list(
                options.map((o) => `{${o.value}, ${mText(o.label)}, ${o.state ?? "null"}}`),
                I1,
            )})`,
        ];
    }
    return [
        `${I1}Source = #table(type table [Value = Int64.Type, Label = text], ${list(
            options.map((o) => `{${o.value}, ${mText(o.label)}}`),
            I1,
        )})`,
    ];
}

function liveChoice(spec: ColumnSpec, settings: BuilderSettings): { steps: string[]; last: string } {
    const kind = spec.attribute.kind;
    const cast = optionMetadataType(kind)!;
    const hasGlobal = kind === "choice" || kind === "multichoice";
    const path = `EntityDefinitions(LogicalName='${spec.entity}')/Attributes(LogicalName='${spec.attribute.logicalName}')/Microsoft.Dynamics.CRM.${cast}`;
    const expand = kind === "boolean" ? "OptionSet" : hasGlobal ? "OptionSet($select=Options),GlobalOptionSet($select=Options)" : "OptionSet($select=Options)";
    const steps = [
        `${I1}Source = Json.Document(Web.Contents(${urlRef(settings)}, [`,
        `${I2}RelativePath = ${mText(path)},`,
        `${I2}Query = [#"$select" = "LogicalName", #"$expand" = ${mText(expand)}],`,
        `${I2}Headers = [Accept = "application/json"]`,
        `${I1}])),`,
        `${I1}GetLabel = (label as nullable record) as nullable text =>`,
        `${I2}if label = null then null`,
        `${I2}else if label[UserLocalizedLabel] <> null then label[UserLocalizedLabel][Label]`,
        `${I2}else if List.IsEmpty(label[LocalizedLabels]) then null`,
        `${I2}else label[LocalizedLabels]{0}[Label],`,
    ];
    if (hasGlobal) {
        steps.push(`${I1}OptionSet = if Record.FieldOrDefault(Source, "OptionSet") <> null then Source[OptionSet] else Source[GlobalOptionSet],`);
    } else {
        steps.push(`${I1}OptionSet = Source[OptionSet],`);
    }
    if (kind === "boolean") {
        steps.push(
            `${I1}Result = #table(type table [Value = logical, Label = text], {`,
            `${I2}{true, GetLabel(OptionSet[TrueOption][Label])},`,
            `${I2}{false, GetLabel(OptionSet[FalseOption][Label])}`,
            `${I1}})`,
        );
        return { steps, last: "Result" };
    }
    if (kind === "status") {
        steps.push(
            `${I1}Options = Table.FromRecords(List.Transform(OptionSet[Options], (o) => [Value = o[Value], Label = GetLabel(o[Label]), State = o[State]]), {"Value", "Label", "State"}),`,
            `${I1}Result = Table.TransformColumnTypes(Options, {{"Value", Int64.Type}, {"Label", type text}, {"State", Int64.Type}})`,
        );
    } else {
        steps.push(
            `${I1}Options = Table.FromRecords(List.Transform(OptionSet[Options], (o) => [Value = o[Value], Label = GetLabel(o[Label])]), {"Value", "Label"}),`,
            `${I1}Result = Table.TransformColumnTypes(Options, {{"Value", Int64.Type}, {"Label", type text}})`,
        );
    }
    return { steps, last: "Result" };
}

/** One query per choice column: Value/Label table to relate to the main table's value column. */
export function generateChoiceQuery(
    spec: ColumnSpec,
    table: TableInfo | undefined,
    settings: BuilderSettings,
    connectionUrl: string,
    staticOptions: ChoiceOption[] | null,
): ChoiceQuery {
    const queryName = choiceQueryName(table, spec);
    const what = `${table?.displayName ?? spec.entity} › ${spec.attribute.displayName} (${spec.entity}.${spec.attribute.logicalName})`;
    let body: string[];
    let last: string;
    let comments: string[];
    if (staticOptions) {
        comments = [`Choice labels for ${what}`, `Static values read when this query was generated; regenerate after the choices change.`];
        body = staticChoice(spec, staticOptions);
        last = "Source";
    } else {
        comments = [`Choice labels for ${what}`, `Read live from Dataverse metadata on every refresh.`];
        const live = liveChoice(spec, settings);
        body = live.steps;
        last = live.last;
    }
    const code = [
        `// Suggested query name: ${queryName}`,
        ...header(comments, staticOptions ? { ...settings, useParameter: false } : settings, connectionUrl),
        "let",
        ...(staticOptions ? [] : urlStep(settings, connectionUrl)),
        ...body,
        "in",
        `${I1}${last}`,
    ].join("\n");
    return { id: spec.id, queryName, entity: spec.entity, attribute: spec.attribute.logicalName, kind: spec.attribute.kind, code };
}
