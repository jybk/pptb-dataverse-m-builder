import type { QueryColumn } from "./mcode";
import { relationshipTmdl, stableId, tmdlName, type ModelQuery, type Relationship } from "./model";
// Power BI Desktop's base theme, as Desktop saves it next to a report (from Microsoft's fabric-cicd samples, MIT)
import baseTheme from "./pbip/CY24SU10.json?raw";

export { baseTheme };

/**
 * A Power BI project (.pbip): a semantic model in TMDL with one table per query, the URL parameter and
 * the relationships, plus an empty report that opens it. Opening the .pbip in Power BI Desktop shows the
 * model without data; Refresh loads it.
 *
 * The files follow what Power BI Desktop itself writes. The report uses the original report.json format,
 * which every Desktop that opens projects reads, and newer versions convert when they save.
 */
export interface ProjectFile {
    /** Relative to the project folder, with forward slashes. */
    path: string;
    content: string;
}

export interface PowerBiProject {
    /** Folder name, which is also the name of the .pbip, report and semantic model. */
    name: string;
    files: ProjectFile[];
}

export interface ProjectTable {
    name: string;
    code: string;
    columns: QueryColumn[];
    kind: "main" | ModelQuery["kind"];
}

export interface PowerBiProjectInput {
    name: string;
    main: { name: string; code: string; columns: QueryColumn[] };
    queries: ModelQuery[];
    relationships: Relationship[];
    /** The URL parameter the queries refer to, or null when the URL is written into them. */
    parameter: { name: string; code: string } | null;
}

const SCHEMA = "https://developer.microsoft.com/json-schemas/fabric";
export const THEME = "CY24SU10";

/** A name that's safe as a file or folder name on Windows and macOS. */
export function fileSafeName(name: string): string {
    const safe = name
        .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "")
        .replace(/[. ]+$/, "")
        .trim();
    return safe || "Dataverse";
}

/** A table's file name the way Power BI Desktop writes them: characters that aren't allowed are %-encoded. */
function tableFileName(name: string): string {
    return name.replace(/[<>:"/\\|?*%\u0000-\u001f]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`) + ".tmdl";
}

/** A TMDL text value: as is, unless it would be read differently (leading or trailing space, a leading quote). */
function tmdlText(value: string): string {
    return /^\s|\s$|^"/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** M lines inside a TMDL property, indented the way Power BI Desktop writes them; empty lines keep the indent too. */
function tmdlExpression(code: string, depth: number): string[] {
    const at = "\t".repeat(depth);
    return code.split(/\r?\n/).map((line) => at + line);
}

/** The query without the "Suggested query name" line: in a project the table already has that name. */
export function withoutNameHint(code: string): string {
    return code.replace(/^\/\/ Suggested query name: .*\r?\n/, "");
}

export interface ColumnShape {
    dataType: string;
    formatString?: string;
    dateOnly?: boolean;
}

/** The model data type for a column the query gives an M type. Untyped columns load as text. */
export function columnShape(mType: string | null): ColumnShape {
    switch (mType) {
        case "Int64.Type":
            return { dataType: "int64", formatString: "0" };
        case "type number":
            return { dataType: "double" };
        case "Currency.Type":
            return { dataType: "decimal" };
        case "type logical":
            return { dataType: "boolean" };
        case "type date":
            return { dataType: "dateTime", formatString: "Long Date", dateOnly: true };
        case "type datetime":
            return { dataType: "dateTime", formatString: "General Date" };
        default:
            return { dataType: "string" };
    }
}

/** Bridge tables only connect two others, so they're hidden from report authors. */
export function isHiddenTable(t: ProjectTable): boolean {
    return t.kind === "bridge";
}

/** Keys of lookup tables only serve the relationships. */
export function isHiddenColumn(t: ProjectTable, c: QueryColumn): boolean {
    return t.kind === "lookup" && c === t.columns[0];
}

/** The main query first, then the Model tab's queries. */
export function projectTables(input: Pick<PowerBiProjectInput, "main" | "queries">): ProjectTable[] {
    return [
        { name: input.main.name, code: input.main.code, columns: input.main.columns, kind: "main" },
        ...input.queries.map((q) => ({ name: q.queryName, code: q.code, columns: q.columns, kind: q.kind })),
    ];
}

function tableTmdl(t: ProjectTable): string {
    const hiddenTable = isHiddenTable(t);
    const lines = [`table ${tmdlName(t.name)}`, ...(hiddenTable ? ["\tisHidden"] : []), `\tlineageTag: ${stableId(`table|${t.name}`)}`, ""];
    for (const c of t.columns) {
        const shape = columnShape(c.mType);
        const hidden = isHiddenColumn(t, c);
        lines.push(
            `\tcolumn ${tmdlName(c.name)}`,
            `\t\tdataType: ${shape.dataType}`,
            ...(hidden ? ["\t\tisHidden"] : []),
            ...(shape.formatString ? [`\t\tformatString: ${shape.formatString}`] : []),
            `\t\tlineageTag: ${stableId(`column|${t.name}|${c.name}`)}`,
            ...(c.webUrl ? ["\t\tdataCategory: WebUrl"] : []),
            `\t\tsummarizeBy: ${c.summarize ? "sum" : "none"}`,
            `\t\tsourceColumn: ${tmdlText(c.name)}`,
            "",
            "\t\tannotation SummarizationSetBy = Automatic",
            ...(shape.dateOnly ? ["", "\t\tannotation UnderlyingDateTimeDataType = Date"] : []),
            ...(shape.dataType === "double" ? ["", '\t\tannotation PBI_FormatHint = {"isGeneralNumber":true}'] : []),
            "",
        );
    }
    lines.push(
        `\tpartition ${tmdlName(t.name)} = m`,
        "\t\tmode: import",
        "\t\tsource =",
        ...tmdlExpression(withoutNameHint(t.code), 4),
        "",
        "\tannotation PBI_ResultType = Table",
        "",
    );
    return lines.join("\n");
}

function json(value: unknown): string {
    return JSON.stringify(value, null, 2) + "\n";
}

/** The report's settings as Power BI Desktop writes them, a JSON string inside the report layout. */
export function reportConfig(): string {
    return JSON.stringify({
        version: "5.59",
        themeCollection: { baseTheme: { name: THEME, version: "5.61", type: 2 } },
        activeSectionIndex: 0,
        defaultDrillFilterOtherVisuals: true,
        linguisticSchemaSyncVersion: 2,
        settings: {
            useNewFilterPaneExperience: true,
            allowChangeFilterTypes: true,
            useStylableVisualContainerHeader: true,
            queryLimitOption: 6,
            useEnhancedTooltips: true,
            exportDataMode: 1,
            useDefaultAggregateDisplayName: true,
        },
    });
}

/** The base theme, which the report refers to and which is saved next to it. */
export const themePackage = {
    resourcePackage: {
        disabled: false,
        items: [{ name: THEME, path: `BaseThemes/${THEME}.json`, type: 202 }],
        name: "SharedResources",
        type: 2,
    },
};

/** An empty report page, in the report.json format Power BI Desktop writes. */
function reportJson(): string {
    return json({
        config: reportConfig(),
        layoutOptimization: 0,
        resourcePackages: [themePackage],
        sections: [
            {
                config: "{}",
                displayName: "Page 1",
                displayOption: 1,
                filters: "[]",
                height: 720.0,
                name: "ReportSection",
                visualContainers: [],
                width: 1280.0,
            },
        ],
    });
}

export function buildPowerBiProject(input: PowerBiProjectInput): PowerBiProject {
    const name = fileSafeName(input.name);
    const report = `${name}.Report`;
    const model = `${name}.SemanticModel`;
    const tables = projectTables(input);
    const queryOrder = [...(input.parameter ? [input.parameter.name] : []), ...tables.map((t) => t.name)];

    const files: ProjectFile[] = [
        {
            path: `${name}.pbip`,
            content: json({
                $schema: `${SCHEMA}/pbip/pbipProperties/1.0.0/schema.json`,
                version: "1.0",
                artifacts: [{ report: { path: report } }],
                settings: { enableAutoRecovery: true },
            }),
        },
        { path: ".gitignore", content: "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n" },
        {
            path: `${report}/definition.pbir`,
            content: json({
                $schema: `${SCHEMA}/item/report/definitionProperties/2.0.0/schema.json`,
                version: "4.0",
                datasetReference: { byPath: { path: `../${model}` } },
            }),
        },
        { path: `${report}/report.json`, content: reportJson() },
        { path: `${report}/StaticResources/SharedResources/BaseThemes/${THEME}.json`, content: baseTheme },
        {
            path: `${model}/definition.pbism`,
            content: json({ $schema: `${SCHEMA}/item/semanticModel/definitionProperties/1.0.0/schema.json`, version: "4.0", settings: {} }),
        },
        { path: `${model}/definition/database.tmdl`, content: "database\n\tcompatibilityLevel: 1550\n\n" },
        {
            path: `${model}/definition/model.tmdl`,
            content: [
                "model Model",
                "\tculture: en-US",
                "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
                "\tsourceQueryCulture: en-US",
                "\tdataAccessOptions",
                "\t\tlegacyRedirects",
                "\t\treturnErrorValuesAsNull",
                "",
                `annotation PBI_QueryOrder = ${JSON.stringify(queryOrder)}`,
                "",
                "annotation __PBI_TimeIntelligenceEnabled = 1",
                "",
                'annotation PBI_ProTooling = ["DevMode"]',
                "",
                ...tables.map((t) => `ref table ${tmdlName(t.name)}`),
                "",
            ].join("\n"),
        },
    ];
    if (input.parameter) {
        files.push({
            path: `${model}/definition/expressions.tmdl`,
            content: [
                `expression ${tmdlName(input.parameter.name)} = ${input.parameter.code}`,
                `\tlineageTag: ${stableId(`expression|${input.parameter.name}`)}`,
                "",
                "\tannotation PBI_ResultType = Text",
                "",
            ].join("\n"),
        });
    }
    if (input.relationships.length > 0) {
        files.push({ path: `${model}/definition/relationships.tmdl`, content: input.relationships.map((r) => relationshipTmdl(r)).join("\n\n") + "\n\n" });
    }
    for (const t of tables) files.push({ path: `${model}/definition/tables/${tableFileName(t.name)}`, content: tableTmdl(t) });
    return { name, files };
}
