import {
    Accordion,
    AccordionHeader,
    AccordionItem,
    AccordionPanel,
    Badge,
    Body1,
    Button,
    Caption1,
    Field,
    Input,
    makeStyles,
    MessageBar,
    MessageBarBody,
    Radio,
    RadioGroup,
    Spinner,
    Switch,
    Subtitle2,
    tokens,
} from "@fluentui/react-components";
import { ArrowRight12Regular, Copy16Regular, DataBarVertical16Regular, DocumentTable16Regular, Save16Regular } from "@fluentui/react-icons";
import type { DataModel, ModelNote, ModelQuery } from "../lib/model";
import type { BuilderSettings } from "../lib/types";
import { CodeView } from "./CodeView";

const useStyles = makeStyles({
    // The tab scrolls as a whole; its sections keep their natural height instead of being squeezed
    root: {
        flex: 1,
        minHeight: 0,
        overflow: "auto",
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalM,
        padding: tokens.spacingHorizontalM,
        "& > *": { flexShrink: 0 },
    },
    row: { display: "flex", flexWrap: "wrap", alignItems: "center", columnGap: tokens.spacingHorizontalL, rowGap: tokens.spacingVerticalXS },
    section: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
    nameField: { maxWidth: "420px" },
    notes: { margin: 0, paddingLeft: "18px", display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
    note: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: tokens.spacingHorizontalS },
    actions: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, flexWrap: "wrap" },
    codeBox: { display: "flex", flexDirection: "column", maxHeight: "420px", minHeight: "120px" },
    header: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, minWidth: 0 },
    // The name and badge stay on one line; only the source wraps
    fixed: { flexShrink: 0, whiteSpace: "nowrap" },
    source: { minWidth: 0 },
    relTable: { borderCollapse: "collapse", fontSize: tokens.fontSizeBase200, width: "100%" },
    th: {
        textAlign: "left",
        padding: "4px 8px",
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        fontWeight: tokens.fontWeightSemibold,
        whiteSpace: "nowrap",
    },
    td: { padding: "4px 8px", borderBottom: `1px solid ${tokens.colorNeutralStroke3}`, verticalAlign: "top" },
    column: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
    steps: { margin: 0, paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "4px" },
    empty: { color: tokens.colorNeutralForeground3 },
});

function ColumnRef(props: { table: string; column: string }) {
    const styles = useStyles();
    return (
        <span>
            {props.table} <span className={styles.column}>[{props.column}]</span>
        </span>
    );
}

/**
 * The Model tab: the queries that go with the main query (choice tables, bridge tables) and the
 * relationships between them, as a TMDL script for Power BI and as a list to follow by hand.
 */
const KIND_BADGE: Record<ModelQuery["kind"], { label: string; color: "brand" | "informative" | "success" }> = {
    choice: { label: "Choice table", color: "informative" },
    bridge: { label: "Bridge", color: "brand" },
    lookup: { label: "Lookup table", color: "success" },
};

export function ModelTab(props: {
    data: DataModel | null;
    /** What the user typed as the main query's name; empty means the suggestion. */
    mainNameInput: string;
    suggestedMainName: string;
    onMainName: (name: string) => void;
    settings: BuilderSettings;
    updateSettings: (patch: Partial<BuilderSettings>) => void;
    choiceLoading: boolean;
    choiceError: string | null;
    onFix: (fix: NonNullable<ModelNote["fix"]>) => void;
    onCopy: (text: string, what: string) => void;
    onSave: (defaultName: string, text: string) => void;
    canExport: boolean;
    exporting: "project" | "template" | "workbook" | null;
    onSaveWorkbook: () => void;
    onSaveProject: () => void;
    onSaveTemplate: () => void;
}) {
    const styles = useStyles();
    const { data } = props;
    if (!data) return <div className={styles.root} />;
    const nothing = data.queries.length === 0;

    return (
        <div className={styles.root}>
            <Body1>
                The tables that go with the main query, and the relationships between them. Save it all as a Power BI project, or add the queries to a file of
                your own and apply the relationships.
            </Body1>

            <Field
                className={styles.nameField}
                label="Main query name"
                hint="The project and the relationships use this name. In a file of your own, name the query from the Power Query (M) tab exactly like this."
            >
                <Input value={props.mainNameInput} placeholder={props.suggestedMainName} onChange={(_, d) => props.onMainName(d.value)} />
            </Field>

            {data.notes.length > 0 && (
                <MessageBar intent="info" layout="multiline">
                    <MessageBarBody>
                        <ul className={styles.notes}>
                            {data.notes.map((n) => (
                                <li key={n.text}>
                                    <span className={styles.note}>
                                        {n.text}
                                        {n.fix && (
                                            <Button size="small" appearance="outline" onClick={() => props.onFix(n.fix!)}>
                                                {n.fix.label}
                                            </Button>
                                        )}

                                    </span>
                                </li>
                            ))}
                        </ul>
                    </MessageBarBody>
                </MessageBar>
            )}

            <div className={styles.section}>
                <Subtitle2>Export</Subtitle2>
                <div className={styles.actions}>
                    <Button appearance="primary" size="small" icon={<DataBarVertical16Regular />} disabled={!props.canExport} onClick={props.onSaveProject}>
                        {props.exporting === "project" ? "Saving…" : "Save Power BI project…"}
                    </Button>
                    <Button size="small" icon={<DataBarVertical16Regular />} disabled={!props.canExport} onClick={props.onSaveTemplate}>
                        {props.exporting === "template" ? "Saving…" : "Save Power BI template…"}
                    </Button>
                    <Button size="small" icon={<DocumentTable16Regular />} disabled={!props.canExport} onClick={props.onSaveWorkbook}>
                        {props.exporting === "workbook" ? "Saving…" : "Save Excel workbook…"}
                    </Button>
                </div>
                <Caption1>
                    The project holds the main query and every query below as tables, with their columns, the URL parameter and the relationships; open its
                    .pbip in Power BI Desktop and choose Refresh. Projects need two preview features turned on. The template holds the same and opens in any
                    Power BI Desktop, the Report Server edition too: it asks for the URL, then loads the data. The workbook holds the main query, loaded to a
                    sheet and refreshed when it opens.
                </Caption1>
            </div>

            <div className={styles.section}>
                <div className={styles.row}>
                    <Subtitle2>Queries</Subtitle2>
                    <span style={{ flex: 1 }} />
                    <Button
                        size="small"
                        icon={<Copy16Regular />}
                        disabled={nothing}
                        onClick={() => props.onCopy(data.queries.map((q) => q.code).join("\n\n"), "All model queries")}
                    >
                        Copy all
                    </Button>
                </div>
                <div className={styles.row}>
                    <Caption1>Choice tables:</Caption1>
                    <RadioGroup
                        layout="horizontal"
                        value={props.settings.choiceTables}
                        onChange={(_, d) => props.updateSettings({ choiceTables: d.value as "live" | "static" })}
                    >
                        <Radio value="live" label="Live (read from metadata on refresh)" />
                        <Radio value="static" label="Static (values embedded now)" />
                    </RadioGroup>
                </div>
                <div className={styles.row}>
                    <Caption1>Lookup tables:</Caption1>
                    <Switch
                        aria-label="Lookup tables"
                        checked={props.settings.lookupTables}
                        onChange={(_, d) => props.updateSettings({ lookupTables: d.checked })}
                        label={props.settings.lookupTables ? "One per table the lookup columns point to, with key and name" : "Off"}
                    />
                </div>
                {props.choiceLoading && <Spinner size="small" label="Reading choice values…" />}
                {props.choiceError && (
                    <MessageBar intent="error">
                        <MessageBarBody>{props.choiceError}</MessageBarBody>
                    </MessageBar>
                )}
                {nothing ? (
                    <Body1 className={styles.empty}>
                        The query has no choice, status, yes/no, multi-select{props.settings.lookupTables ? " or lookup" : ""} columns
                        {props.settings.lookupTables ? "" : " (lookup tables are off)"}, so there's nothing to relate to it yet.
                    </Body1>
                ) : (
                    <Accordion multiple collapsible>
                        {data.queries.map((q) => (
                            <AccordionItem key={q.id} value={q.id}>
                                <AccordionHeader size="small">
                                    <span className={styles.header}>
                                        <span className={styles.fixed}>{q.queryName}</span>
                                        <Badge className={styles.fixed} size="small" appearance="tint" color={KIND_BADGE[q.kind].color}>
                                            {KIND_BADGE[q.kind].label}
                                        </Badge>
                                        <Caption1 className={styles.source}>{q.source}</Caption1>
                                    </span>
                                </AccordionHeader>
                                <AccordionPanel>
                                    <div className={styles.actions} style={{ marginBottom: 6 }}>
                                        <Button size="small" icon={<Copy16Regular />} onClick={() => props.onCopy(q.code, q.queryName)}>
                                            Copy
                                        </Button>
                                        <Button size="small" icon={<Save16Regular />} onClick={() => props.onSave(`${q.queryName}.pq`, q.code)}>
                                            Save .pq
                                        </Button>
                                    </div>
                                    <div className={styles.codeBox}>
                                        <CodeView code={q.code} />
                                    </div>
                                </AccordionPanel>
                            </AccordionItem>
                        ))}
                    </Accordion>
                )}
            </div>

            {data.relationships.length > 0 && (
                <div className={styles.section}>
                    <Subtitle2>Relationships</Subtitle2>
                    <table className={styles.relTable}>
                        <thead>
                            <tr>
                                <th className={styles.th}>From (many)</th>
                                <th className={styles.th} aria-label="to" />
                                <th className={styles.th}>To (one)</th>
                                <th className={styles.th}>Filter</th>
                            </tr>
                        </thead>
                        <tbody>
                            {data.relationships.map((r) => (
                                <tr key={r.id}>
                                    <td className={styles.td}>
                                        <ColumnRef {...r.from} />
                                    </td>
                                    <td className={styles.td}>
                                        <ArrowRight12Regular />
                                    </td>
                                    <td className={styles.td}>
                                        <ColumnRef {...r.to} />
                                    </td>
                                    <td className={styles.td}>
                                        {r.bothDirections ? "Both directions" : "Single"}
                                        {r.active ? "" : " · inactive"}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                    {data.tmdl && (
                        <>
                            <div className={styles.row}>
                                <Caption1>TMDL script for Power BI Desktop</Caption1>
                                <span style={{ flex: 1 }} />
                                <div className={styles.actions}>
                                    <Button
                                        size="small"
                                        appearance="primary"
                                        icon={<Copy16Regular />}
                                        onClick={() => props.onCopy(data.tmdl!, "Relationship script")}
                                    >
                                        Copy script
                                    </Button>
                                    <Button
                                        size="small"
                                        icon={<Save16Regular />}
                                        onClick={() => props.onSave(`${data.mainName} relationships.tmdl`, data.tmdl!)}
                                    >
                                        Save .tmdl
                                    </Button>
                                </div>
                            </div>
                            <div className={styles.codeBox}>
                                <CodeView code={data.tmdl} language="tmdl" />
                            </div>
                        </>
                    )}
                    <ol className={styles.steps}>
                        <li>
                            Add each query above in Power Query (<b>Get data → Blank query → Advanced Editor</b>), named as shown, and name the main query{" "}
                            <b>{data.mainName}</b>.
                        </li>
                        <li>
                            Before loading, untick <b>Autodetect new relationships after data is loaded</b> (File → Options and settings → Options → Current
                            file → Data Load). Otherwise Power BI may add relationships of its own between columns that only share a name, such as the
                            choice tables' Value columns. If it already did, delete those in <b>Manage relationships</b>.
                        </li>
                        <li>
                            Choose <b>Close &amp; Apply</b>, so the tables are loaded.
                        </li>
                        <li>
                            Open <b>TMDL view</b> on the left side of Power BI Desktop, paste the script in a tab and click <b>Apply</b>. Applying it again
                            replaces these relationships rather than adding copies.
                        </li>
                        <li>In Excel's data model or elsewhere, create the relationships by hand from the list above.</li>
                    </ol>
                </div>
            )}
        </div>
    );
}
