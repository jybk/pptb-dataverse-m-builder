import { Button, Caption1, Checkbox, Input, makeStyles, mergeClasses, Select, Spinner, Subtitle2, Switch, tokens, Tooltip } from "@fluentui/react-components";
import { Search16Regular } from "@fluentui/react-icons";
import { memo, useMemo, useState, type ReactNode } from "react";
import { defaultBaseName, defaultModeFor, modeOptionsFor } from "../lib/columns";
import { kindLabel } from "../lib/metadata";
import type { AttributeInfo, BuilderSettings, ColumnOverride, ColumnOverrides, ColumnSpec, OutputColumn } from "../lib/types";

const useStyles = makeStyles({
    root: { display: "flex", flexDirection: "column", minHeight: 0, height: "100%", minWidth: 0 },
    head: {
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalXS,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    titleRow: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap", minHeight: "32px" },
    toolRow: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
    scroll: { flex: 1, overflow: "auto", minHeight: 0 },
    table: { width: "100%", borderCollapse: "collapse", fontSize: tokens.fontSizeBase200 },
    th: {
        position: "sticky",
        top: 0,
        zIndex: 1,
        textAlign: "left",
        fontWeight: tokens.fontWeightSemibold,
        padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
        backgroundColor: tokens.colorNeutralBackground3,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        whiteSpace: "nowrap",
    },
    td: { padding: `2px ${tokens.spacingHorizontalS}`, borderBottom: `1px solid ${tokens.colorNeutralStroke3}`, verticalAlign: "middle" },
    rowIncluded: { backgroundColor: tokens.colorNeutralBackground1Selected },
    colName: { display: "flex", flexDirection: "column", minWidth: "140px" },
    meta: { fontSize: "11px", lineHeight: "15px" },
    display: { fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground1 },
    logical: { color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace, fontSize: "11px" },
    type: { color: tokens.colorNeutralForeground2 },
    unknown: { color: tokens.colorPaletteDarkOrangeForeground1 },
    outputs: { color: tokens.colorNeutralForeground3, fontSize: "11px", marginTop: "2px" },
    empty: { padding: tokens.spacingHorizontalL, color: tokens.colorNeutralForeground3 },
    outputCell: { display: "flex", flexWrap: "wrap", gap: "4px", alignItems: "center" },
    modeSelect: { flex: "0 0 132px", minWidth: 0 },
    nameInput: { flex: "1 1 120px", minWidth: "100px", "& input": { minWidth: 0, width: "100%" } },
});

export interface GridRow {
    id: string;
    attribute: AttributeInfo;
    included: boolean;
    spec: ColumnSpec | null;
    /** Extra context shown under the column, e.g. "via contact1" or an alias. */
    note?: string;
}

interface Props {
    title: ReactNode;
    subtitle?: ReactNode;
    headerActions?: ReactNode;
    rows: GridRow[];
    loading?: boolean;
    emptyText: string;
    includeLabel: string;
    outputsByCol: Map<string, OutputColumn[]>;
    settings: BuilderSettings;
    overrides: ColumnOverrides;
    onToggle: (ids: string[], included: boolean) => void;
    onOverride: (id: string, patch: ColumnOverride) => void;
}

export function ColumnGrid(props: Props) {
    const styles = useStyles();
    const [search, setSearch] = useState("");
    const [onlyIncluded, setOnlyIncluded] = useState(false);

    const visible = useMemo(() => {
        const q = search.trim().toLowerCase();
        return props.rows.filter((r) => {
            if (onlyIncluded && !r.included) return false;
            if (!q) return true;
            return r.attribute.displayName.toLowerCase().includes(q) || r.attribute.logicalName.includes(q) || (r.note ?? "").toLowerCase().includes(q);
        });
    }, [props.rows, search, onlyIncluded]);

    const includedCount = props.rows.filter((r) => r.included).length;
    const allVisibleIncluded = visible.length > 0 && visible.every((r) => r.included);

    return (
        <div className={styles.root}>
            <div className={styles.head}>
                <div className={styles.titleRow}>
                    <Subtitle2>{props.title}</Subtitle2>
                    {props.subtitle}
                    <span style={{ flex: 1 }} />
                    {props.headerActions}
                </div>
                <div className={styles.toolRow}>
                    <Input
                        size="small"
                        contentBefore={<Search16Regular />}
                        placeholder="Search columns"
                        value={search}
                        onChange={(_, d) => setSearch(d.value)}
                        aria-label="Search columns"
                        style={{ minWidth: 180, flex: "1 1 180px", maxWidth: 320 }}
                    />
                    <Switch label={`${props.includeLabel} only`} checked={onlyIncluded} onChange={(_, d) => setOnlyIncluded(d.checked)} />
                    <Caption1>
                        {includedCount} of {props.rows.length} {props.includeLabel.toLowerCase()}
                    </Caption1>
                    <span style={{ flex: 1 }} />
                    <Button
                        size="small"
                        appearance="subtle"
                        disabled={visible.length === 0}
                        onClick={() =>
                            props.onToggle(
                                visible.map((r) => r.id),
                                !allVisibleIncluded,
                            )
                        }
                    >
                        {allVisibleIncluded ? "Clear shown" : "Select shown"}
                    </Button>
                </div>
            </div>
            <div className={styles.scroll}>
                {props.loading ? (
                    <Spinner size="small" label="Loading columns…" style={{ padding: 24 }} />
                ) : visible.length === 0 ? (
                    <div className={styles.empty}>{props.rows.length === 0 ? props.emptyText : "No columns match."}</div>
                ) : (
                    <table className={styles.table}>
                        <thead>
                            <tr>
                                <th className={styles.th} style={{ width: 32 }} aria-label={props.includeLabel} />
                                <th className={styles.th}>Column</th>
                                <th className={styles.th} style={{ width: "55%" }}>
                                    Output &amp; name
                                </th>
                            </tr>
                        </thead>
                        <tbody>
                            {visible.map((row) => (
                                <GridRowView
                                    key={row.id}
                                    row={row}
                                    outputs={props.outputsByCol.get(row.id)}
                                    settings={props.settings}
                                    override={props.overrides[row.id]}
                                    onToggle={props.onToggle}
                                    onOverride={props.onOverride}
                                />
                            ))}
                        </tbody>
                    </table>
                )}
            </div>
        </div>
    );
}

const GridRowView = memo(function GridRowView(props: {
    row: GridRow;
    outputs: OutputColumn[] | undefined;
    settings: BuilderSettings;
    override: ColumnOverride | undefined;
    onToggle: (ids: string[], included: boolean) => void;
    onOverride: (id: string, patch: ColumnOverride) => void;
}) {
    const styles = useStyles();
    const { row, override, settings } = props;
    const a = row.attribute;
    const modes = row.spec ? modeOptionsFor(row.spec) : null;
    const currentMode = override?.mode ?? (row.spec ? defaultModeFor(row.spec, settings) : null) ?? "";
    const placeholder = row.spec ? defaultBaseName(row.spec, settings.naming) : a.displayName;
    const outputNames = (props.outputs ?? []).map((o) => o.name);
    const showOutputs = outputNames.length > 1 || (outputNames.length === 1 && outputNames[0] !== (override?.name?.trim() || placeholder));

    return (
        <tr className={mergeClasses(row.included && styles.rowIncluded)}>
            <td className={styles.td}>
                <Checkbox checked={row.included} onChange={(_, d) => props.onToggle([row.id], !!d.checked)} aria-label={`Include ${a.displayName}`} />
            </td>
            <td className={styles.td}>
                <div className={styles.colName}>
                    <span className={styles.display}>{a.displayName}</span>
                    <span className={styles.meta}>
                        <span className={styles.logical}>{a.logicalName}</span>
                        {row.note ? <span className={styles.logical}>{` · ${row.note}`}</span> : null}
                        <span className={mergeClasses(styles.type, a.kind === "unknown" && styles.unknown)}>
                            {" · "}
                            {a.kind === "unknown" ? (
                                <Tooltip content="Not found in metadata: the column is passed through without a type" relationship="description">
                                    <span>Unknown type</span>
                                </Tooltip>
                            ) : (
                                kindLabel(a)
                            )}
                        </span>
                    </span>
                </div>
            </td>
            <td className={styles.td}>
                {row.included ? (
                    <>
                        <div className={styles.outputCell}>
                            {modes ? (
                                <Select
                                    size="small"
                                    className={styles.modeSelect}
                                    value={currentMode}
                                    onChange={(_, d) => props.onOverride(row.id, { mode: d.value })}
                                    aria-label={`Output for ${a.displayName}`}
                                >
                                    {modes.map((m) => (
                                        <option key={m.value} value={m.value}>
                                            {m.label}
                                        </option>
                                    ))}
                                </Select>
                            ) : null}
                            <Input
                                size="small"
                                className={styles.nameInput}
                                value={override?.name ?? ""}
                                placeholder={placeholder}
                                onChange={(_, d) => props.onOverride(row.id, { name: d.value })}
                                aria-label={`Name for ${a.displayName}`}
                            />
                        </div>
                        {showOutputs && <div className={styles.outputs}>→ {outputNames.join(" · ")}</div>}
                    </>
                ) : null}
            </td>
        </tr>
    );
});
