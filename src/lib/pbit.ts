import JSZip from "jszip";
import type { QueryColumn } from "./mcode";
import { stableId, type Relationship } from "./model";
import {
    baseTheme,
    columnShape,
    isHiddenColumn,
    isHiddenTable,
    projectTables,
    reportConfig,
    themePackage,
    THEME,
    withoutNameHint,
    type PowerBiProjectInput,
    type ProjectTable,
} from "./pbip";

/**
 * A Power BI template (.pbit): the same model as the Power BI project, in the file format every Power BI
 * Desktop opens, the Report Server edition included, without preview features. Opening it asks for the
 * URL parameter, then loads the data into a new report.
 *
 * Microsoft doesn't document the parts of a template. These follow what Power BI Desktop itself writes,
 * in the February 2025 release that the project files follow too: a zip (an Open Packaging Conventions
 * package) whose JSON parts are UTF-16 without a byte order mark, and whose model is the JSON form of the
 * semantic model that Microsoft's Analysis Services libraries read and write.
 */

/** Power BI Desktop release the parts are modelled on. */
const DESKTOP = { version: "2.140.679.0", release: "2025.02", reportVersion: "5.59", themeVersion: "5.61" };
const TEMPLATE_VERSION = "1.30";
const COMPATIBILITY_LEVEL = 1550;

/** UTF-16 little-endian without a byte order mark, as Desktop writes a template's parts. */
export function utf16(text: string): Uint8Array {
    const bytes = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        bytes[i * 2] = code & 0xff;
        bytes[i * 2 + 1] = code >> 8;
    }
    return bytes;
}

/** M split into lines, as Desktop stores expressions in a template's model. */
function lines(code: string): string[] {
    return code.split(/\r?\n/);
}

function columnJson(t: ProjectTable, c: QueryColumn) {
    const shape = columnShape(c.mType);
    const annotations = [
        { name: "SummarizationSetBy", value: "Automatic" },
        ...(shape.dateOnly ? [{ name: "UnderlyingDateTimeDataType", value: "Date" }] : []),
        ...(shape.dataType === "double" ? [{ name: "PBI_FormatHint", value: '{"isGeneralNumber":true}' }] : []),
    ];
    return {
        name: c.name,
        dataType: shape.dataType,
        ...(isHiddenColumn(t, c) ? { isHidden: true } : {}),
        ...(shape.formatString ? { formatString: shape.formatString } : {}),
        sourceColumn: c.name,
        lineageTag: stableId(`column|${t.name}|${c.name}`),
        ...(c.webUrl ? { dataCategory: "WebUrl" } : {}),
        summarizeBy: c.summarize ? "sum" : "none",
        annotations,
    };
}

function tableJson(t: ProjectTable) {
    return {
        name: t.name,
        ...(isHiddenTable(t) ? { isHidden: true } : {}),
        lineageTag: stableId(`table|${t.name}`),
        columns: t.columns.map((c) => columnJson(t, c)),
        partitions: [{ name: t.name, mode: "import", source: { type: "m", expression: lines(withoutNameHint(t.code)) } }],
        annotations: [{ name: "PBI_ResultType", value: "Table" }],
    };
}

function relationshipJson(r: Relationship) {
    return {
        name: r.id,
        fromTable: r.from.table,
        fromColumn: r.from.column,
        toTable: r.to.table,
        toColumn: r.to.column,
        ...(r.bothDirections ? { crossFilteringBehavior: "bothDirections" } : {}),
        ...(r.active ? {} : { isActive: false }),
    };
}

/** The template's DataModelSchema: the model as JSON (TMSL), the format Microsoft's libraries read. */
export function templateModel(input: PowerBiProjectInput): object {
    const tables = projectTables(input);
    const queryOrder = [...(input.parameter ? [input.parameter.name] : []), ...tables.map((t) => t.name)];
    return {
        name: stableId(`database|${input.name}`),
        compatibilityLevel: COMPATIBILITY_LEVEL,
        model: {
            culture: "en-US",
            dataAccessOptions: { legacyRedirects: true, returnErrorValuesAsNull: true },
            defaultPowerBIDataSourceVersion: "powerBI_V3",
            sourceQueryCulture: "en-US",
            tables: tables.map(tableJson),
            relationships: input.relationships.map(relationshipJson),
            ...(input.parameter
                ? {
                      expressions: [
                          {
                              name: input.parameter.name,
                              kind: "m",
                              expression: input.parameter.code,
                              lineageTag: stableId(`expression|${input.parameter.name}`),
                              annotations: [{ name: "PBI_ResultType", value: "Text" }],
                          },
                      ],
                  }
                : {}),
            annotations: [
                { name: "PBI_QueryOrder", value: JSON.stringify(queryOrder) },
                { name: "__PBI_TimeIntelligenceEnabled", value: "1" },
                { name: "PBIDesktopVersion", value: `${DESKTOP.version} (${DESKTOP.release.slice(2)})` },
            ],
        },
    };
}

/** The tables laid out in Model view: the main table first, the others in a grid beside it. */
function diagramLayout(tables: ProjectTable[]): object {
    return {
        version: "1.1.0",
        diagrams: [
            {
                ordinal: 0,
                scrollPosition: { x: 0, y: 0 },
                nodes: tables.map((t, i) => ({
                    location: i === 0 ? { x: 0, y: 0 } : { x: 320 * (1 + Math.floor((i - 1) / 3)), y: 340 * ((i - 1) % 3) },
                    nodeIndex: t.name,
                    nodeLineageTag: stableId(`table|${t.name}`),
                    size: { height: 300, width: 234 },
                    zIndex: i,
                })),
                name: "All tables",
                zoomValue: 100,
                pinKeyFieldsToTop: false,
                showExtraHeaderInfo: false,
                hideKeyFieldsWhenCollapsed: false,
                tablesLocked: false,
            },
        ],
        selectedDiagram: "All tables",
        defaultDiagram: "All tables",
    };
}

/** One empty page, in the layout format Desktop writes into templates. */
function reportLayout(): object {
    return {
        id: 0,
        resourcePackages: [themePackage],
        sections: [
            {
                id: 0,
                name: "ReportSection",
                displayName: "Page 1",
                filters: "[]",
                ordinal: 0,
                visualContainers: [],
                config: "{}",
                displayOption: 1,
                width: 1280,
                height: 720,
            },
        ],
        config: reportConfig(),
        layoutOptimization: 0,
    };
}

const CONTENT_TYPES =
    '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="json" ContentType="" />' +
    '<Override PartName="/Version" ContentType="" />' +
    '<Override PartName="/DataModelSchema" ContentType="" />' +
    '<Override PartName="/DiagramLayout" ContentType="" />' +
    '<Override PartName="/Report/Layout" ContentType="" />' +
    '<Override PartName="/Settings" ContentType="application/json" />' +
    '<Override PartName="/Metadata" ContentType="application/json" />' +
    "</Types>";

/** The template's parts, in the order Desktop writes them, as bytes. */
export function templateParts(input: PowerBiProjectInput): Array<[string, Uint8Array]> {
    const tables = projectTables(input);
    const jsonPart = (value: unknown) => utf16(JSON.stringify(value));
    return [
        ["Version", utf16(TEMPLATE_VERSION)],
        ["[Content_Types].xml", new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(CONTENT_TYPES)])],
        ["DataModelSchema", jsonPart(templateModel(input))],
        ["DiagramLayout", jsonPart(diagramLayout(tables))],
        ["Report/Layout", jsonPart(reportLayout())],
        [
            "Settings",
            jsonPart({ Version: 3, ReportSettings: {}, QueriesSettings: { TypeDetectionEnabled: true, RelationshipImportEnabled: true, Version: DESKTOP.version } }),
        ],
        [
            "Metadata",
            jsonPart({
                Version: 5,
                AutoCreatedRelationships: [],
                FileDescription: `${input.main.name} and its related tables, from Dataverse M Builder`,
                CreatedFrom: "Cloud",
                CreatedFromRelease: DESKTOP.release,
            }),
        ],
        [`Report/StaticResources/SharedResources/BaseThemes/${THEME}.json`, new TextEncoder().encode(baseTheme)],
    ];
}

export async function buildPowerBiTemplate(input: PowerBiProjectInput): Promise<Uint8Array> {
    const zip = new JSZip();
    // A fixed date, so the same model gives the same file
    const date = new Date(Date.UTC(2025, 1, 1));
    // Only file entries, as in Desktop's templates: package readers can reject folder entries
    for (const [path, bytes] of templateParts(input)) zip.file(path, bytes, { date, binary: true, createFolders: false });
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    return withDeflateVersions(bytes);
}

/**
 * Sets the zip version fields the way Desktop's templates have them: 2.0 needed to extract (the version
 * for deflated entries; the zip library writes 1.0), made by 4.5. Lenient readers ignore these fields,
 * but Desktop's package reader is strict, so the template matches Desktop's own byte for byte here.
 */
export function withDeflateVersions(bytes: Uint8Array): Uint8Array {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const end = bytes.length - 22; // the end-of-central-directory record, without a comment
    if (view.getUint32(end, true) !== 0x06054b50) throw new Error("Unexpected zip layout");
    const count = view.getUint16(end + 10, true);
    let at = view.getUint32(end + 16, true);
    for (let i = 0; i < count; i++) {
        if (view.getUint32(at, true) !== 0x02014b50) throw new Error("Unexpected zip layout");
        view.setUint16(at + 4, 45, true);
        view.setUint16(at + 6, 20, true);
        const local = view.getUint32(at + 42, true);
        if (view.getUint32(local, true) !== 0x04034b50) throw new Error("Unexpected zip layout");
        view.setUint16(local + 4, 20, true);
        at += 46 + view.getUint16(at + 28, true) + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    }
    return bytes;
}
