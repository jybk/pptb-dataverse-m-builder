import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { detectTimeZone } from "../lib/timezone";
import { DEFAULT_SETTINGS, type BuilderSettings } from "../lib/types";

type Connection = ToolBoxAPI.Connection;

export function hasToolbox(): boolean {
    return typeof window !== "undefined" && !!window.toolboxAPI && !!window.dataverseAPI;
}

/** Active connection, refreshed when ToolBox reports connection changes. */
export function useConnection() {
    const [connection, setConnection] = useState<Connection | null>(null);
    const [loading, setLoading] = useState(true);

    const refresh = useCallback(async () => {
        if (!hasToolbox()) {
            setLoading(false);
            return;
        }
        try {
            setConnection(await window.toolboxAPI.connections.getActiveConnection());
        } catch (err) {
            console.error("Couldn't read the active connection", err);
            setConnection(null);
        } finally {
            setLoading(false);
        }
    }, []);

    useEffect(() => {
        void refresh();
    }, [refresh]);

    return { connection, loading, refresh };
}

/** Subscribe to ToolBox events for the lifetime of the component. */
export function useToolboxEvents(handler: (event: string, data: unknown) => void) {
    const ref = useRef(handler);
    ref.current = handler;
    useEffect(() => {
        if (!hasToolbox()) return;
        const cb = (_e: unknown, payload: ToolBoxAPI.ToolBoxEventPayload) => ref.current(payload?.event, payload?.data);
        window.toolboxAPI.events.on(cb);
        return () => {
            try {
                window.toolboxAPI.events.off?.(cb);
            } catch {
                /* older hosts can't unsubscribe */
            }
        };
    }, []);
}

export function useTheme() {
    const [theme, setTheme] = useState<"light" | "dark">("light");
    const refresh = useCallback(async () => {
        if (!hasToolbox()) {
            setTheme(window.matchMedia?.("(prefers-color-scheme: dark)").matches ? "dark" : "light");
            return;
        }
        try {
            setTheme((await window.toolboxAPI.utils.getCurrentTheme()) === "dark" ? "dark" : "light");
        } catch {
            /* keep current theme */
        }
    }, []);
    useEffect(() => {
        void refresh();
    }, [refresh]);
    return { theme, refreshTheme: refresh };
}

const SETTINGS_KEY = "builderSettings";

/**
 * Saved settings on top of the defaults. Settings saved before daylight saving existed keep their
 * fixed offset if local time was in use; everyone else starts from this computer's time zone.
 */
export function withTimeZone<T extends Partial<BuilderSettings>>(saved: T, detect: () => Pick<BuilderSettings, "utcOffsetMinutes" | "daylightSaving"> = detectTimeZone): BuilderSettings & T {
    if (saved.daylightSaving !== undefined || saved.dateMode === "offset") return { ...DEFAULT_SETTINGS, ...saved };
    let zone = {};
    try {
        zone = detect();
    } catch {
        /* keep the defaults */
    }
    return { ...DEFAULT_SETTINGS, ...saved, ...zone };
}

/** What is stored: the settings, plus the model-driven app chosen for record links in each environment. */
type StoredSettings = BuilderSettings & { appIdByEnv?: Record<string, string> };

/**
 * The app id that applies to an environment. An app id belongs to one environment, so it is kept per
 * environment; a value saved before that (one global id) is treated as belonging to the first
 * environment that asks.
 */
export function appIdFor(stored: StoredSettings, environment: string | null): string {
    if (!environment) return "";
    const byEnv = stored.appIdByEnv;
    if (byEnv) return byEnv[environment] ?? "";
    return stored.appId ?? "";
}

/** Builder settings, persisted with the ToolBox settings API. `environment` is the connected environment's URL. */
export function useBuilderSettings(environment: string | null) {
    const [stored, setSettings] = useState<StoredSettings>(DEFAULT_SETTINGS);
    const loaded = useRef(false);

    useEffect(() => {
        (async () => {
            if (hasToolbox()) {
                try {
                    const saved = await window.toolboxAPI.settings.get(SETTINGS_KEY);
                    setSettings(withTimeZone(saved && typeof saved === "object" ? (saved as Partial<StoredSettings>) : {}));
                } catch (err) {
                    console.warn("Couldn't read saved settings", err);
                }
            }
            loaded.current = true;
        })();
    }, []);

    useEffect(() => {
        if (!loaded.current || !hasToolbox()) return;
        const t = window.setTimeout(() => {
            window.toolboxAPI.settings.set(SETTINGS_KEY, stored).catch((err: unknown) => console.warn("Couldn't save settings", err));
        }, 400);
        return () => window.clearTimeout(t);
    }, [stored]);

    // Pin a pre-existing global app id to the first environment seen, so it stops following the user to other environments
    useEffect(() => {
        if (!loaded.current || !environment) return;
        setSettings((s) => (s.appIdByEnv ? s : { ...s, appId: "", appIdByEnv: s.appId ? { [environment]: s.appId } : {} }));
    }, [environment, stored]);

    const settings = useMemo<BuilderSettings>(() => {
        const { appIdByEnv: _byEnv, ...rest } = stored;
        return { ...rest, appId: appIdFor(stored, environment) };
    }, [stored, environment]);

    const update = useCallback(
        (patch: Partial<BuilderSettings>) =>
            setSettings((s) => {
                if (patch.appId === undefined) return { ...s, ...patch };
                const { appId, ...rest } = patch;
                const byEnv = { ...(s.appIdByEnv ?? {}) };
                if (environment) {
                    if (appId) byEnv[environment] = appId;
                    else delete byEnv[environment];
                }
                return { ...s, ...rest, appId: "", appIdByEnv: byEnv };
            }),
        [environment],
    );
    // Reset keeps the per-environment app choices out: they are settings too
    return { settings, update, reset: () => setSettings(withTimeZone({})) };
}

export async function notify(title: string, body: string, type: "info" | "success" | "warning" | "error" = "info") {
    if (!hasToolbox()) {
        console.log(`[${type}] ${title}: ${body}`);
        return;
    }
    try {
        await window.toolboxAPI.utils.showNotification({ title, body, type, duration: type === "error" ? 6000 : 3000 });
    } catch (err) {
        console.warn("Notification failed", err);
    }
}

export async function copyText(text: string, what: string) {
    try {
        if (hasToolbox()) await window.toolboxAPI.utils.copyToClipboard(text);
        else await navigator.clipboard.writeText(text);
        await notify("Copied", `${what} copied to the clipboard.`, "success");
    } catch (err) {
        await notify("Couldn't copy", err instanceof Error ? err.message : String(err), "error");
    }
}

export async function saveText(defaultName: string, text: string) {
    try {
        const path = await window.toolboxAPI.fileSystem.saveFile(defaultName, text, [
            { name: "Power Query", extensions: ["pq", "m"] },
            { name: "Text", extensions: ["txt"] },
        ]);
        if (path) await notify("Saved", path, "success");
    } catch (err) {
        await notify("Couldn't save the file", err instanceof Error ? err.message : String(err), "error");
    }
}
