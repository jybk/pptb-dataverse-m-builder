import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DebouncedSaver } from "../lib/saver";
import { detectTimeZone } from "../lib/timezone";
import { DEFAULT_SETTINGS, type BuilderSettings } from "../lib/types";

type Connection = ToolBoxAPI.Connection;

/**
 * Saves `value` under `key` with the ToolBox settings API shortly after it last changed. A save still
 * waiting is written when the tool is hidden or closed, and when `key` changes or becomes null
 * (for example while another environment's session loads).
 */
export function useDebouncedSave(key: string | null, value: unknown, delay: number) {
    const saver = useMemo(() => new DebouncedSaver((k, v) => window.toolboxAPI.settings.set(k, v), delay), [delay]);

    useEffect(() => {
        if (key && hasToolbox()) saver.schedule(key, value);
        else saver.flush();
    }, [saver, key, value]);

    useEffect(() => {
        const flush = () => saver.flush();
        const onVisibility = () => {
            if (document.visibilityState === "hidden") saver.flush();
        };
        window.addEventListener("pagehide", flush);
        document.addEventListener("visibilitychange", onVisibility);
        return () => {
            window.removeEventListener("pagehide", flush);
            document.removeEventListener("visibilitychange", onVisibility);
            saver.flush();
        };
    }, [saver]);
}

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

    useDebouncedSave(loaded.current ? SETTINGS_KEY : null, stored, 400);

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

/** Saves bytes (a workbook, say) through ToolBox's save dialog, which writes them as they are. */
export async function saveBytes(defaultName: string, bytes: Uint8Array, filters: Array<{ name: string; extensions: string[] }>) {
    try {
        const path = await window.toolboxAPI.fileSystem.saveFile(defaultName, bytes, filters);
        if (path) await notify("Saved", path, "success");
    } catch (err) {
        await notify("Couldn't save the file", err instanceof Error ? err.message : String(err), "error");
    }
}

async function exists(path: string): Promise<boolean> {
    try {
        await window.toolboxAPI.fileSystem.stat(path);
        return true;
    } catch {
        return false;
    }
}

/**
 * Writes a project's files into a new folder inside the one the user picks, "Name (2)" if Name is taken,
 * so nothing that's there is overwritten. Picking the folder gives the tool access to write inside it.
 */
export async function saveProjectFolder(project: { name: string; files: Array<{ path: string; content: string }> }, what: string, open: string) {
    try {
        const parent = await window.toolboxAPI.fileSystem.selectPath({ type: "folder", title: `Choose where to save the ${what}`, buttonLabel: "Save here" });
        if (!parent) return;
        const sep = parent.includes("\\") && !parent.includes("/") ? "\\" : "/";
        const base = parent.replace(/[\\/]+$/, "") + sep + project.name;
        let folder = base;
        for (let n = 2; await exists(folder); n++) folder = `${base} (${n})`;
        const made = new Set<string>();
        for (const file of project.files) {
            const path = folder + sep + file.path.split("/").join(sep);
            const dir = path.slice(0, path.lastIndexOf(sep));
            if (!made.has(dir)) {
                await window.toolboxAPI.fileSystem.createDirectory(dir);
                made.add(dir);
            }
            await window.toolboxAPI.fileSystem.writeText(path, file.content);
        }
        await notify("Saved", `Open ${open} in ${folder}`, "success");
    } catch (err) {
        await notify(`Couldn't save the ${what}`, err instanceof Error ? err.message : String(err), "error");
    }
}
