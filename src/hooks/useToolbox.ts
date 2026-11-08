import { useCallback, useEffect, useRef, useState } from "react";
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

/** Builder settings, persisted with the ToolBox settings API. */
export function useBuilderSettings() {
    const [settings, setSettings] = useState<BuilderSettings>(DEFAULT_SETTINGS);
    const loaded = useRef(false);

    useEffect(() => {
        (async () => {
            if (hasToolbox()) {
                try {
                    const saved = await window.toolboxAPI.settings.get(SETTINGS_KEY);
                    if (saved && typeof saved === "object") setSettings({ ...DEFAULT_SETTINGS, ...(saved as Partial<BuilderSettings>) });
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
            window.toolboxAPI.settings.set(SETTINGS_KEY, settings).catch((err: unknown) => console.warn("Couldn't save settings", err));
        }, 400);
        return () => window.clearTimeout(t);
    }, [settings]);

    const update = useCallback((patch: Partial<BuilderSettings>) => setSettings((s) => ({ ...s, ...patch })), []);
    return { settings, update, reset: () => setSettings(DEFAULT_SETTINGS) };
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
