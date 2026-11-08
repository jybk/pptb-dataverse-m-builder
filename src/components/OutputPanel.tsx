import {
    Accordion,
    AccordionHeader,
    AccordionItem,
    AccordionPanel,
    Body1,
    Button,
    Caption1,
    Label,
    makeStyles,
    MessageBar,
    MessageBarBody,
    MessageBarTitle,
    Radio,
    RadioGroup,
    Select,
    Spinner,
    Switch,
    Tab,
    TabList,
    tokens,
    ToggleButton,
    Tooltip,
} from "@fluentui/react-components";
import { ArrowReply16Regular, Copy16Regular, Edit16Regular, Play16Regular, Save16Regular } from "@fluentui/react-icons";
import { useState, type ReactNode } from "react";
import type { ChoiceQuery } from "../lib/mcode";
import type { PreviewResult } from "../lib/preview";
import type { BuilderSettings, QueryModel } from "../lib/types";
import { CodeView } from "./CodeView";

const useStyles = makeStyles({
    root: { display: "flex", flexDirection: "column", minHeight: 0, height: "100%", minWidth: 0 },
    tabs: { borderBottom: `1px solid ${tokens.colorNeutralStroke2}`, paddingLeft: tokens.spacingHorizontalS, flexShrink: 0 },
    body: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM, overflow: "hidden" },
    scrollBody: { flex: 1, minHeight: 0, overflow: "auto", display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
    options: { display: "flex", flexWrap: "wrap", alignItems: "center", columnGap: tokens.spacingHorizontalL, rowGap: tokens.spacingVerticalXS },
    field: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS },
    actions: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, flexWrap: "wrap" },
    warnings: { margin: 0, paddingLeft: "18px" },
    center: { margin: "auto", textAlign: "center", color: tokens.colorNeutralForeground3, maxWidth: "380px" },
    steps: { margin: 0, paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "4px" },
    previewWrap: { flex: 1, minHeight: 0, overflow: "auto", border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium },
    previewTable: { borderCollapse: "collapse", fontSize: tokens.fontSizeBase200, minWidth: "100%" },
    pth: {
        position: "sticky",
        top: 0,
        backgroundColor: tokens.colorNeutralBackground3,
        textAlign: "left",
        padding: "4px 8px",
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        whiteSpace: "nowrap",
        fontWeight: tokens.fontWeightSemibold,
    },
    ptd: { padding: "3px 8px", borderBottom: `1px solid ${tokens.colorNeutralStroke3}`, whiteSpace: "nowrap", maxWidth: "320px", overflow: "hidden", textOverflow: "ellipsis" },
    nullCell: { color: tokens.colorNeutralForeground4, fontStyle: "italic" },
    codeBox: { display: "flex", flexDirection: "column", maxHeight: "420px", minHeight: "120px" },
});

export type OutputTab = "m" | "choices" | "parameter" | "fetchxml" | "preview";

export interface PreviewState {
    loading: boolean;
    error: string | null;
    result: PreviewResult | null;
    ranAt: number | null;
    /** The M code the preview was run for, to flag a stale preview. */
    ranFor?: string | null;
}

/** UTC−12:00 … UTC+14:00 in half hours, plus the quarter-hour zones (Nepal, Eucla, Chatham). */
const OFFSETS = [...Array.from({ length: 53 }, (_, i) => -720 + i * 30), 345, 525, 765].sort((a, b) => a - b);

function offsetLabel(minutes: number): string {
    const sign = minutes < 0 ? "−" : "+";
    const abs = Math.abs(minutes);
    return `UTC${sign}${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
}

export function OutputPanel(props: {
    tab: OutputTab;
    onTab: (tab: OutputTab) => void;
    mode: "table" | "fetchxml";
    model: QueryModel | null;
    modelError: string | null;
    columnCount: number;
    mCode: string | null;
    settings: BuilderSettings;
    updateSettings: (patch: Partial<BuilderSettings>) => void;
    parameterCode: string;
    choiceQueries: ChoiceQuery[];
    choiceLoading: boolean;
    choiceError: string | null;
    onEditAsFetchXml: () => void;
    preview: PreviewState;
    onRunPreview: (rows: number) => void;
    canReturn: boolean;
    onReturn: () => void;
    onCopy: (text: string, what: string) => void;
    onSave: (defaultName: string, text: string) => void;
    queryName: string;
    emptyHint: ReactNode;
}) {
    const styles = useStyles();
    const { settings, model } = props;
    const [previewRows, setPreviewRows] = useState(50);
    const odataAvailable = !!model?.odata;

    const tabs = (
        <TabList className={styles.tabs} size="small" selectedValue={props.tab} onTabSelect={(_, d) => props.onTab(d.value as OutputTab)}>
            <Tab value="m">Power Query (M)</Tab>
            <Tab value="choices">Choice tables{props.choiceQueries.length ? ` (${props.choiceQueries.length})` : ""}</Tab>
            <Tab value="parameter">Parameter</Tab>
            <Tab value="fetchxml">FetchXML</Tab>
            <Tab value="preview">Preview</Tab>
        </TabList>
    );

    if (!model) {
        return (
            <div className={styles.root}>
                {tabs}
                <div className={styles.body}>
                    {props.modelError ? (
                        <MessageBar intent="error">
                            <MessageBarBody>{props.modelError}</MessageBarBody>
                        </MessageBar>
                    ) : (
                        <div className={styles.center}>{props.emptyHint}</div>
                    )}
                </div>
            </div>
        );
    }

    return (
        <div className={styles.root}>
            {tabs}
            {props.tab === "m" && (
                <div className={styles.body}>
                    <div className={styles.options}>
                        <div className={styles.field}>
                            <Label size="small">Query</Label>
                            <ToggleButton
                                size="small"
                                checked={settings.queryStyle === "fetchxml" || !odataAvailable}
                                onClick={() => props.updateSettings({ queryStyle: "fetchxml" })}
                            >
                                FetchXML
                            </ToggleButton>
                            <Tooltip
                                content={
                                    odataAvailable
                                        ? "OData $select query for a single table"
                                        : "OData is only available for a single table picked on the left (no joins or aggregates)"
                                }
                                relationship="description"
                            >
                                <ToggleButton
                                    size="small"
                                    disabled={!odataAvailable}
                                    checked={settings.queryStyle === "odata" && odataAvailable}
                                    onClick={() => props.updateSettings({ queryStyle: "odata" })}
                                >
                                    OData
                                </ToggleButton>
                            </Tooltip>
                        </div>
                        {props.mode === "table" && (
                            <Switch label="Active rows only" checked={settings.activeOnly} onChange={(_, d) => props.updateSettings({ activeOnly: d.checked })} />
                        )}
                        <Switch label="Record URL" checked={settings.includeRecordUrl} onChange={(_, d) => props.updateSettings({ includeRecordUrl: d.checked })} />
                        <div className={styles.field}>
                            <Label size="small" htmlFor="dmb-names">
                                Names
                            </Label>
                            <Select id="dmb-names" size="small" value={settings.naming} onChange={(_, d) => props.updateSettings({ naming: d.value as BuilderSettings["naming"] })}>
                                <option value="display">Display names</option>
                                <option value="logical">Logical names</option>
                                <option value="schema">Schema names</option>
                            </Select>
                        </div>
                        <div className={styles.field}>
                            <Label size="small" htmlFor="dmb-dates">
                                Date/time
                            </Label>
                            <Select
                                id="dmb-dates"
                                size="small"
                                value={settings.dateMode}
                                onChange={(_, d) => props.updateSettings({ dateMode: d.value as BuilderSettings["dateMode"] })}
                            >
                                <option value="utc">UTC</option>
                                <option value="offset">Fixed offset</option>
                                <option value="formatted">Formatted text</option>
                            </Select>
                            {settings.dateMode === "offset" && (
                                <Select
                                    size="small"
                                    aria-label="UTC offset"
                                    value={String(settings.utcOffsetMinutes)}
                                    onChange={(_, d) => props.updateSettings({ utcOffsetMinutes: Number(d.value) })}
                                >
                                    {OFFSETS.map((m) => (
                                        <option key={m} value={m}>
                                            {offsetLabel(m)}
                                        </option>
                                    ))}
                                </Select>
                            )}
                        </div>
                    </div>
                    {model.warnings.length > 0 && (
                        <MessageBar intent="warning" layout="multiline">
                            <MessageBarBody>
                                <MessageBarTitle>Check before you use this query</MessageBarTitle>
                                <ul className={styles.warnings}>
                                    {model.warnings.map((w, i) => (
                                        <li key={i}>{w}</li>
                                    ))}
                                </ul>
                            </MessageBarBody>
                        </MessageBar>
                    )}
                    <div className={styles.actions}>
                        <Button
                            size="small"
                            appearance="primary"
                            icon={<Copy16Regular />}
                            disabled={!props.mCode}
                            onClick={() => props.onCopy(props.mCode ?? "", "Power Query (M)")}
                        >
                            Copy M
                        </Button>
                        <Button size="small" icon={<Save16Regular />} disabled={!props.mCode} onClick={() => props.onSave(`${props.queryName}.pq`, props.mCode ?? "")}>
                            Save .pq
                        </Button>
                        {props.canReturn && (
                            <Tooltip content="Returns this M query and its FetchXML to the tool that opened this one, and closes this tab" relationship="description">
                                <Button size="small" icon={<ArrowReply16Regular />} onClick={props.onReturn}>
                                    Send back to caller
                                </Button>
                            </Tooltip>
                        )}
                        <span style={{ flex: 1 }} />
                        <Caption1>
                            {props.columnCount} output column{props.columnCount === 1 ? "" : "s"}
                        </Caption1>
                    </div>
                    {props.mCode && <CodeView code={props.mCode} />}
                    <Accordion collapsible>
                        <AccordionItem value="how">
                            <AccordionHeader size="small">How to use this in Power BI or Excel</AccordionHeader>
                            <AccordionPanel>
                                <ol className={styles.steps}>
                                    <li>
                                        {settings.useParameter ? (
                                            <>
                                                Create a text parameter named <b>{settings.parameterName || "DataverseUrl"}</b> (see the Parameter tab). One parameter serves every
                                                query, so the Power BI service only needs one credential.
                                            </>
                                        ) : (
                                            <>The environment URL is written into the query (change this under Settings).</>
                                        )}
                                    </li>
                                    <li>
                                        <b>Get data → Blank query</b>, open <b>Advanced Editor</b>, paste the code and click <b>Done</b>.
                                    </li>
                                    <li>
                                        When Power BI asks for credentials for <b>…/api/data/{settings.apiVersion || "v9.2"}</b>, choose <b>Organizational account</b> and sign in.
                                        For scheduled refresh, set the same OAuth2 credential on the data source in the Power BI service; if its connection test fails, tick{" "}
                                        <b>Skip test connection</b>.
                                    </li>
                                    <li>Rename the query, then add the choice tables if you want to relate labels to values.</li>
                                </ol>
                            </AccordionPanel>
                        </AccordionItem>
                    </Accordion>
                </div>
            )}

            {props.tab === "choices" && (
                <div className={styles.scrollBody}>
                    <div className={styles.options}>
                        <RadioGroup layout="horizontal" value={settings.choiceTables} onChange={(_, d) => props.updateSettings({ choiceTables: d.value as "live" | "static" })}>
                            <Radio value="live" label="Live (read from metadata on refresh)" />
                            <Radio value="static" label="Static (values embedded now)" />
                        </RadioGroup>
                        <span style={{ flex: 1 }} />
                        <Button
                            size="small"
                            icon={<Copy16Regular />}
                            disabled={props.choiceQueries.length === 0}
                            onClick={() => props.onCopy(props.choiceQueries.map((q) => q.code).join("\n\n"), "All choice queries")}
                        >
                            Copy all
                        </Button>
                    </div>
                    <Body1>
                        One query per choice, status or yes/no column, with a Value and Label column. Set the column to <b>Value</b> (or <b>Label + value</b>) on the left to relate
                        it to these tables in your model.
                    </Body1>
                    {props.choiceLoading && <Spinner size="small" label="Reading choice values…" />}
                    {props.choiceError && (
                        <MessageBar intent="error">
                            <MessageBarBody>{props.choiceError}</MessageBarBody>
                        </MessageBar>
                    )}
                    {props.choiceQueries.length === 0 && !props.choiceLoading && <div className={styles.center}>The query has no choice, status or yes/no columns.</div>}
                    <Accordion multiple collapsible>
                        {props.choiceQueries.map((q) => (
                            <AccordionItem key={q.id} value={q.id}>
                                <AccordionHeader size="small">
                                    {q.queryName} <Caption1 style={{ marginLeft: 8 }}>{`${q.entity}.${q.attribute}`}</Caption1>
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
                </div>
            )}

            {props.tab === "parameter" && (
                <div className={styles.scrollBody}>
                    <Body1>
                        The queries read the environment's Web API address from a parameter, so you can point a report at another environment, or move to another API version, by
                        changing one value.
                    </Body1>
                    <ol className={styles.steps}>
                        <li>
                            In Power Query, choose <b>Manage Parameters → New Parameter</b>.
                        </li>
                        <li>
                            Name it <b>{settings.parameterName || "DataverseUrl"}</b>, set Type to <b>Text</b> and paste the URL below as the current value (no trailing slash).
                        </li>
                    </ol>
                    <Body1>Or create a blank query with that name and paste this into the Advanced Editor:</Body1>
                    <div className={styles.actions}>
                        <Button size="small" icon={<Copy16Regular />} onClick={() => props.onCopy(props.parameterCode, "Parameter")}>
                            Copy
                        </Button>
                    </div>
                    <div className={styles.codeBox} style={{ minHeight: 60 }}>
                        <CodeView code={props.parameterCode} />
                    </div>
                    {!settings.useParameter && (
                        <MessageBar intent="info">
                            <MessageBarBody>The generated queries currently embed the URL instead. Turn the parameter back on under Settings.</MessageBarBody>
                        </MessageBar>
                    )}
                </div>
            )}

            {props.tab === "fetchxml" && (
                <div className={styles.body}>
                    <div className={styles.actions}>
                        <Button size="small" icon={<Copy16Regular />} onClick={() => props.onCopy(model.fetchXml, "FetchXML")}>
                            Copy
                        </Button>
                        <Button size="small" icon={<Save16Regular />} onClick={() => props.onSave(`${props.queryName}.xml`, model.fetchXml)}>
                            Save
                        </Button>
                        {props.mode === "table" && (
                            <Button size="small" icon={<Edit16Regular />} onClick={props.onEditAsFetchXml}>
                                Edit as FetchXML
                            </Button>
                        )}
                        <span style={{ flex: 1 }} />
                        <Caption1>As embedded in the M query (paging is added per request)</Caption1>
                    </div>
                    <CodeView code={model.fetchXml} language="xml" />
                </div>
            )}

            {props.tab === "preview" && (
                <div className={styles.body}>
                    <div className={styles.actions}>
                        <Button
                            size="small"
                            appearance="primary"
                            icon={<Play16Regular />}
                            disabled={props.preview.loading || props.columnCount === 0}
                            onClick={() => props.onRunPreview(previewRows)}
                        >
                            Run preview
                        </Button>
                        <Select size="small" aria-label="Rows" value={String(previewRows)} onChange={(_, d) => setPreviewRows(Number(d.value))}>
                            {[10, 50, 100, 250].map((n) => (
                                <option key={n} value={n}>
                                    First {n} rows
                                </option>
                            ))}
                        </Select>
                        <span style={{ flex: 1 }} />
                        {props.preview.result && <Caption1>{props.preview.result.rows.length} rows · shaped like the M output</Caption1>}
                    </div>
                    {props.preview.loading && <Spinner size="small" label="Running query…" />}
                    {props.preview.error && (
                        <MessageBar intent="error" layout="multiline">
                            <MessageBarBody>
                                <MessageBarTitle>The query failed</MessageBarTitle>
                                {props.preview.error}
                            </MessageBarBody>
                        </MessageBar>
                    )}
                    {props.preview.result && props.preview.ranFor !== undefined && props.preview.ranFor !== props.mCode && (
                        <MessageBar intent="warning">
                            <MessageBarBody>The query changed since this preview. Run it again to see the new shape.</MessageBarBody>
                        </MessageBar>
                    )}
                    {props.preview.result && props.preview.result.unusedKeys.length > 0 && (
                        <MessageBar intent="info">
                            <MessageBarBody>Returned but not used in the M output: {props.preview.result.unusedKeys.join(", ")}</MessageBarBody>
                        </MessageBar>
                    )}
                    {!props.preview.result && !props.preview.loading && !props.preview.error && (
                        <div className={styles.center}>Runs the query against the connected environment and shows the rows the way the M query will shape them.</div>
                    )}
                    {props.preview.result && (
                        <div className={styles.previewWrap}>
                            <table className={styles.previewTable}>
                                <thead>
                                    <tr>
                                        {props.preview.result.columns.map((c) => (
                                            <th key={c} className={styles.pth}>
                                                {c}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {props.preview.result.rows.map((r, i) => (
                                        <tr key={i}>
                                            {r.map((v, j) => (
                                                <td key={j} className={styles.ptd} title={v ?? "null"}>
                                                    {v === null ? <span className={styles.nullCell}>null</span> : v}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}
