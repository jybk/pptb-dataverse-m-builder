import {
    Button,
    Combobox,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    Field,
    Input,
    makeStyles,
    Option,
    Select,
    SpinButton,
    Switch,
    tokens,
} from "@fluentui/react-components";
import { useEffect, useState } from "react";
import { CHOICE_MODES, LOOKUP_MODES } from "../lib/columns";
import { loadApps, type ModelDrivenApp } from "../lib/dataverse";
import type { BuilderSettings } from "../lib/types";

const useStyles = makeStyles({
    grid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}` },
    full: { gridColumn: "1 / -1" },
    section: { gridColumn: "1 / -1", fontWeight: tokens.fontWeightSemibold, marginTop: tokens.spacingVerticalS, color: tokens.colorNeutralForeground2 },
});

export function SettingsDialog(props: {
    open: boolean;
    onClose: () => void;
    settings: BuilderSettings;
    update: (patch: Partial<BuilderSettings>) => void;
    reset: () => void;
    connected: boolean;
    /** The connected environment; the list of apps belongs to it. */
    environment: string | null;
}) {
    const styles = useStyles();
    const { settings, update } = props;
    const [apps, setApps] = useState<ModelDrivenApp[] | null>(null);

    useEffect(() => setApps(null), [props.environment]);

    useEffect(() => {
        if (!props.open || !props.connected || apps) return;
        loadApps()
            .then(setApps)
            .catch(() => setApps([]));
    }, [props.open, props.connected, apps]);

    const selectedApp = apps?.find((a) => a.id === settings.appId);

    return (
        <Dialog open={props.open} onOpenChange={(_, d) => !d.open && props.onClose()}>
            <DialogSurface style={{ maxWidth: 640 }}>
                <DialogBody>
                    <DialogTitle>Settings</DialogTitle>
                    <DialogContent>
                        <div className={styles.grid}>
                            <div className={styles.section}>Connection in the generated queries</div>
                            <Field
                                label="Use a parameter for the environment URL"
                                className={styles.full}
                                hint="Recommended: one data source credential in the Power BI service, and easy switching between environments."
                            >
                                <Switch checked={settings.useParameter} onChange={(_, d) => update({ useParameter: d.checked })} />
                            </Field>
                            <Field
                                label="Parameter name"
                                validationMessage={!/^[A-Za-z_][A-Za-z0-9_ ]*$/.test(settings.parameterName) ? "Use letters, digits, spaces or underscores" : undefined}
                            >
                                <Input value={settings.parameterName} disabled={!settings.useParameter} onChange={(_, d) => update({ parameterName: d.value })} />
                            </Field>
                            <Field label="Web API version" hint="Goes into the parameter value, e.g. …/api/data/v9.2">
                                <Select value={settings.apiVersion} onChange={(_, d) => update({ apiVersion: d.value })}>
                                    {["v9.2", "v9.1", "v9.0"].map((v) => (
                                        <option key={v}>{v}</option>
                                    ))}
                                </Select>
                            </Field>
                            <Field label="Rows per request" hint="Dataverse returns at most 5,000 rows per page.">
                                <SpinButton
                                    value={settings.pageSize}
                                    min={100}
                                    max={5000}
                                    step={500}
                                    onChange={(_, d) => {
                                        const v = d.value ?? Number(d.displayValue);
                                        if (Number.isFinite(v)) update({ pageSize: Math.min(5000, Math.max(100, Math.round(v as number))) });
                                    }}
                                />
                            </Field>

                            <div className={styles.section}>Default output per column type</div>
                            <Field label="Lookups">
                                <Select value={settings.lookupMode} onChange={(_, d) => update({ lookupMode: d.value as BuilderSettings["lookupMode"] })}>
                                    {LOOKUP_MODES.map((m) => (
                                        <option key={m.value} value={m.value}>
                                            {m.label}
                                        </option>
                                    ))}
                                </Select>
                            </Field>
                            <Field label="Choices and status">
                                <Select value={settings.choiceMode} onChange={(_, d) => update({ choiceMode: d.value as BuilderSettings["choiceMode"] })}>
                                    {CHOICE_MODES.map((m) => (
                                        <option key={m.value} value={m.value}>
                                            {m.label}
                                        </option>
                                    ))}
                                </Select>
                            </Field>
                            <Field label="Yes/No columns">
                                <Select value={settings.booleanMode} onChange={(_, d) => update({ booleanMode: d.value as BuilderSettings["booleanMode"] })}>
                                    {CHOICE_MODES.map((m) => (
                                        <option key={m.value} value={m.value}>
                                            {m.label === "Value" ? "Value (true/false)" : m.label}
                                        </option>
                                    ))}
                                </Select>
                            </Field>

                            <div className={styles.section}>Record URL</div>
                            <Field label="Open records in app" className={styles.full} hint="Optional. Without an app, links open in the user's default app.">
                                <Combobox
                                    placeholder={props.connected ? (apps ? "Default app" : "Loading apps…") : "Connect to list apps"}
                                    value={selectedApp?.name ?? (settings.appId ? settings.appId : "")}
                                    selectedOptions={settings.appId ? [settings.appId] : []}
                                    onOptionSelect={(_, d) => update({ appId: d.optionValue ?? "" })}
                                    freeform={false}
                                    clearable
                                >
                                    {(apps ?? []).map((a) => (
                                        <Option key={a.id} value={a.id} text={a.name}>
                                            {a.name}
                                        </Option>
                                    ))}
                                </Combobox>
                            </Field>
                        </div>
                    </DialogContent>
                    <DialogActions>
                        <Button appearance="secondary" onClick={props.reset}>
                            Reset to defaults
                        </Button>
                        <Button appearance="primary" onClick={props.onClose}>
                            Done
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
}
