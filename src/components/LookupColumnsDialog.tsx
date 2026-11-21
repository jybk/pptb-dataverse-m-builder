import { Button, Caption1, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, makeStyles } from "@fluentui/react-components";
import { useMemo } from "react";
import type { LookupTableQuery } from "../lib/model";
import type { AttributeInfo, BuilderSettings, ColumnOverride, LookupColumns, OutputColumn, TableInfo } from "../lib/types";
import { ColumnGrid, type GridRow } from "./ColumnGrid";

const useStyles = makeStyles({
    surface: { width: "min(980px, 94vw)", maxWidth: "none" },
    grid: { height: "min(560px, 62vh)", display: "flex", flexDirection: "column", minHeight: 0 },
});

/**
 * Picks the columns of a lookup table beyond its key and name, with the same grid as the main query:
 * tick a column to add it, choose how it comes out and rename it. Key and name are always there.
 */
export function LookupColumnsDialog(props: {
    target: TableInfo;
    queryName: string;
    /** The target table's columns; undefined while they load. */
    attributes: AttributeInfo[] | undefined;
    choice: LookupColumns;
    /** The lookup table as it would be generated with this choice, for the column specs and output names. */
    preview: LookupTableQuery | null;
    settings: BuilderSettings;
    onChange: (choice: LookupColumns) => void;
    onClose: () => void;
}) {
    const styles = useStyles();
    const { target, choice, preview } = props;
    const fixed = useMemo(() => new Set([`|${target.primaryIdAttribute}`, ...(target.primaryNameAttribute ? [`|${target.primaryNameAttribute}`] : [])]), [target]);

    const rows = useMemo<GridRow[]>(() => {
        const chosen = new Set(choice.columns);
        const specs = new Map((preview?.model.columns ?? []).map((c) => [c.id, c]));
        const sorted = [...(props.attributes ?? [])].sort((a, b) => {
            // Key and name first, as in the lookup table
            const rank = (x: AttributeInfo) => (x.logicalName === target.primaryIdAttribute ? 0 : x.logicalName === target.primaryNameAttribute ? 1 : 2);
            return rank(a) - rank(b) || a.displayName.localeCompare(b.displayName);
        });
        return sorted.map((a) => {
            const id = `|${a.logicalName}`;
            const locked = fixed.has(id);
            const included = locked || chosen.has(a.logicalName);
            return { id, attribute: a, included, locked, spec: included ? (specs.get(id) ?? null) : null, note: locked ? "always included" : undefined };
        });
    }, [props.attributes, choice.columns, preview, fixed, target]);

    const outputsByCol = useMemo(() => {
        const m = new Map<string, OutputColumn[]>();
        for (const o of preview?.outputs ?? []) m.set(o.colId, [...(m.get(o.colId) ?? []), o]);
        return m;
    }, [preview]);

    const toggle = (ids: string[], included: boolean) => {
        const names = ids.filter((id) => !fixed.has(id)).map((id) => id.replace(/^\|/, ""));
        const current = new Set(choice.columns);
        for (const n of names) {
            if (included) current.add(n);
            else current.delete(n);
        }
        // In the table's column order, so the lookup table's columns don't jump around
        const order = new Map((props.attributes ?? []).map((a, i) => [a.logicalName, i]));
        props.onChange({ ...choice, columns: [...current].sort((a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0)) });
    };

    const override = (id: string, patch: ColumnOverride) => {
        if (fixed.has(id)) return;
        props.onChange({ ...choice, overrides: { ...choice.overrides, [id]: { ...(choice.overrides[id] ?? {}), ...patch } } });
    };

    return (
        <Dialog open modalType="modal" onOpenChange={(_, d) => !d.open && props.onClose()}>
            <DialogSurface className={styles.surface}>
                <DialogBody>
                    <DialogTitle>Columns of {props.queryName}</DialogTitle>
                    <DialogContent>
                        <Caption1 block style={{ marginBottom: 8 }}>
                            {props.queryName} reads every row of {target.displayName} ({target.logicalName}). Its key and name are always there; tick other columns to add
                            them, and choose how each comes out, as for the main query. Every column you add is read for every row.
                        </Caption1>
                        <div className={styles.grid}>
                            <ColumnGrid
                                title={target.displayName}
                                subtitle={<Caption1>{target.logicalName}</Caption1>}
                                rows={rows}
                                loading={!props.attributes}
                                emptyText="This table has no readable columns."
                                includeLabel="In lookup table"
                                outputsByCol={outputsByCol}
                                settings={props.settings}
                                overrides={preview?.overrides ?? choice.overrides}
                                onToggle={toggle}
                                onOverride={override}
                            />
                        </div>
                    </DialogContent>
                    <DialogActions>
                        <Button appearance="primary" onClick={props.onClose}>
                            Done
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
}
