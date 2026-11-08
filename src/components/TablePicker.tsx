import { Badge, Button, Caption1, Input, makeStyles, mergeClasses, Select, Spinner, tokens, Tooltip } from "@fluentui/react-components";
import { ArrowClockwise16Regular, Search16Regular } from "@fluentui/react-icons";
import { useMemo, useState } from "react";
import type { TableInfo } from "../lib/types";

const useStyles = makeStyles({
    root: {
        display: "flex",
        flexDirection: "column",
        minHeight: 0,
        height: "100%",
        borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
        backgroundColor: tokens.colorNeutralBackground2,
    },
    head: {
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalXS,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    filters: { display: "flex", gap: tokens.spacingHorizontalXS, alignItems: "center" },
    list: { flex: 1, overflowY: "auto", padding: `${tokens.spacingVerticalXS} 0` },
    item: {
        display: "flex",
        flexDirection: "column",
        width: "100%",
        textAlign: "left",
        border: "none",
        background: "none",
        cursor: "pointer",
        padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
        color: tokens.colorNeutralForeground1,
        fontFamily: tokens.fontFamilyBase,
        borderLeft: "3px solid transparent",
        ":hover": { backgroundColor: tokens.colorSubtleBackgroundHover },
        ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "-2px" },
    },
    selected: {
        backgroundColor: tokens.colorSubtleBackgroundSelected,
        borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
    },
    name: { fontSize: tokens.fontSizeBase300, lineHeight: tokens.lineHeightBase300, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
    logical: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
    empty: { padding: tokens.spacingHorizontalM, color: tokens.colorNeutralForeground3 },
});

type Filter = "all" | "custom" | "standard";

export function TablePicker(props: { tables: TableInfo[] | null; loading: boolean; selected: string | null; onSelect: (logicalName: string) => void; onReload: () => void }) {
    const styles = useStyles();
    const [search, setSearch] = useState("");
    const [filter, setFilter] = useState<Filter>("all");

    const visible = useMemo(() => {
        const q = search.trim().toLowerCase();
        return (props.tables ?? []).filter((t) => {
            if (filter === "custom" && !t.isCustom) return false;
            if (filter === "standard" && t.isCustom) return false;
            return !q || t.displayName.toLowerCase().includes(q) || t.logicalName.includes(q);
        });
    }, [props.tables, search, filter]);

    return (
        <div className={styles.root}>
            <div className={styles.head}>
                <Input
                    size="small"
                    contentBefore={<Search16Regular />}
                    placeholder="Search tables"
                    value={search}
                    onChange={(_, d) => setSearch(d.value)}
                    aria-label="Search tables"
                />
                <div className={styles.filters}>
                    <Select size="small" value={filter} onChange={(_, d) => setFilter(d.value as Filter)} aria-label="Filter tables" style={{ flex: 1, minWidth: 0 }}>
                        <option value="all">All tables</option>
                        <option value="standard">Standard</option>
                        <option value="custom">Custom</option>
                    </Select>
                    <Tooltip content="Reload tables" relationship="label">
                        <Button size="small" appearance="subtle" icon={<ArrowClockwise16Regular />} onClick={props.onReload} />
                    </Tooltip>
                </div>
                {props.tables && (
                    <Caption1>
                        {visible.length} of {props.tables.length} tables
                    </Caption1>
                )}
            </div>
            <div className={styles.list} role="listbox" aria-label="Tables">
                {props.loading && <Spinner size="small" label="Loading tables…" style={{ padding: 16 }} />}
                {!props.loading && visible.length === 0 && <div className={styles.empty}>No tables match.</div>}
                {visible.map((t) => (
                    <button
                        key={t.logicalName}
                        role="option"
                        aria-selected={props.selected === t.logicalName}
                        className={mergeClasses(styles.item, props.selected === t.logicalName && styles.selected)}
                        onClick={() => props.onSelect(t.logicalName)}
                        title={`${t.displayName} (${t.logicalName})`}
                    >
                        <span className={styles.name}>
                            {t.displayName}{" "}
                            {t.isIntersect && (
                                <Badge size="small" appearance="tint" color="informative">
                                    N:N
                                </Badge>
                            )}
                        </span>
                        <span className={styles.logical}>{t.logicalName}</span>
                    </button>
                ))}
            </div>
        </div>
    );
}
