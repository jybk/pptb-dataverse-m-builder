import {
    Badge,
    Button,
    Caption1,
    FluentProvider,
    makeStyles,
    MessageBar,
    MessageBarBody,
    Spinner,
    Tab,
    TabList,
    Title3,
    tokens,
    Tooltip,
    webDarkTheme,
    webLightTheme,
} from "@fluentui/react-components";
import { ArrowClockwise16Regular, Settings20Regular } from "@fluentui/react-icons";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ColumnGrid, type GridRow } from "./components/ColumnGrid";
import { FetchXmlEditor } from "./components/FetchXmlEditor";
import { OutputPanel, type OutputTab, type PreviewState } from "./components/OutputPanel";
import { SettingsDialog } from "./components/SettingsDialog";
import { TablePicker } from "./components/TablePicker";
import { ViewMenu } from "./components/ViewMenu";
import { copyText, hasToolbox, notify, saveText, useBuilderSettings, useConnection, useDebouncedSave, useTheme, useToolboxEvents } from "./hooks/useToolbox";
import { outputWarnings, resolveOutputColumns } from "./lib/columns";
import { errorMessage, loadAttributes, loadOptions, loadTables, loadViews, runFetchPreview, runODataPreview, type ViewInfo } from "./lib/dataverse";
import { buildQueryModel, buildTableFetchXml, fetchEntities, FetchXmlError, linesToXml, parseFetchXml, previewFetchXml, serializeLines } from "./lib/fetchxml";
import { callerExpectsReply, otherTools } from "./lib/invocation";
import { EMPTY_SESSION, MAX_FETCH_LENGTH, parseSession, serializeSession, sessionKey } from "./lib/session";
import { choiceColumns, generateChoiceQuery, generateFetchXmlQuery, generateODataQuery, generateParameterQuery, normaliseUrl, odataSelect, type ChoiceQuery } from "./lib/mcode";
import type { ChoiceOption } from "./lib/metadata";
import { buildPreview } from "./lib/preview";
import type { AttributeInfo, ColumnOverride, ColumnOverrides, OutputColumn, QueryModel, TableInfo } from "./lib/types";

/** The left pane's width, in percent of the window. */
const SPLIT_MIN = 30;
const SPLIT_MAX = 75;

const useStyles = makeStyles({
    root: { display: "flex", flexDirection: "column", height: "100vh", backgroundColor: tokens.colorNeutralBackground1, color: tokens.colorNeutralForeground1, overflow: "hidden" },
    header: {
        display: "flex",
        alignItems: "center",
        gap: tokens.spacingHorizontalM,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        flexWrap: "wrap",
        flexShrink: 0,
    },
    titleBlock: { display: "flex", flexDirection: "column", lineHeight: 1.1 },
    main: { flex: 1, display: "flex", minHeight: 0 },
    left: { display: "flex", minWidth: 0, minHeight: 0 },
    tablePane: { width: "clamp(170px, 18vw, 240px)", flexShrink: 0, minHeight: 0 },
    gridPane: { flex: 1, minWidth: 0, minHeight: 0 },
    fetchPane: { display: "flex", flexDirection: "column", flex: 1, minWidth: 0, minHeight: 0 },
    editorPane: { height: "42%", minHeight: "160px", borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
    editorPaneCollapsed: { flexShrink: 0, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
    splitter: {
        width: "5px",
        cursor: "col-resize",
        flexShrink: 0,
        backgroundColor: tokens.colorNeutralStroke2,
        ":hover": { backgroundColor: tokens.colorBrandStroke1 },
        ":focus-visible": { backgroundColor: tokens.colorBrandStroke1, outline: `2px solid ${tokens.colorStrokeFocus2}` },
    },
    right: { flex: 1, minWidth: "360px", minHeight: 0, display: "flex", flexDirection: "column" },
    center: {
        margin: "auto",
        textAlign: "center",
        maxWidth: "460px",
        color: tokens.colorNeutralForeground2,
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalM,
        padding: tokens.spacingHorizontalXXL,
    },
});

function useDebounced<T>(value: T, ms: number): T {
    const [v, setV] = useState(value);
    useEffect(() => {
        const t = window.setTimeout(() => setV(value), ms);
        return () => window.clearTimeout(t);
    }, [value, ms]);
    return v;
}

/** Pretty-print FetchXML; returns the input unchanged when it doesn't parse (the editor then shows the error). */
function formatFetchXml(xml: string): string {
    try {
        return linesToXml(serializeLines(parseFetchXml(xml).documentElement));
    } catch {
        return xml;
    }
}

const ENV_COLORS: Record<string, "danger" | "warning" | "success" | "informative"> = { Production: "danger", UAT: "warning", Test: "warning", Dev: "success" };

export default function App() {
    const styles = useStyles();
    const { theme, refreshTheme } = useTheme();
    const { connection, loading: connectionLoading, refresh: refreshConnection } = useConnection();
    // The environment's address identifies what is remembered for it (session, record-link app)
    const envKey = connection?.url ? normaliseUrl(connection.url).toLowerCase() : null;
    const { settings, update, reset } = useBuilderSettings(envKey);

    const [tables, setTables] = useState<TableInfo[] | null>(null);
    const [tablesLoading, setTablesLoading] = useState(false);
    const [tablesError, setTablesError] = useState<string | null>(null);
    const [attrCache, setAttrCache] = useState<Map<string, AttributeInfo[]>>(new Map());
    const [attrFailed, setAttrFailed] = useState<Record<string, string>>({});
    /** Tables whose date formats couldn't be read, so date-only columns can't be told apart. */
    const [dateFormatsMissing, setDateFormatsMissing] = useState<Set<string>>(new Set());
    const inflight = useRef(new Set<string>());
    /**
     * Bumped whenever the metadata is reset (connection switch, reload). A load started before that
     * belongs to the previous environment, so its result is dropped instead of landing in the new caches.
     */
    const metadataGeneration = useRef(0);

    const [mode, setMode] = useState<"table" | "fetchxml">("table");
    const [selectedTable, setSelectedTable] = useState<string | null>(null);
    const [selection, setSelection] = useState<Record<string, string[]>>({});
    const [fetchText, setFetchText] = useState("");
    const debouncedFetch = useDebounced(fetchText, 300);
    /** The view the FetchXML came from, and whether it was changed since. */
    const [viewName, setViewName] = useState<string | null>(null);
    const [viewEdited, setViewEdited] = useState(false);
    /** The FetchXML editor is collapsed once a query is loaded; it opens by itself when it's empty or has an error. */
    const [editorOpen, setEditorOpen] = useState(false);
    const [overridesByCtx, setOverridesByCtx] = useState<Record<string, ColumnOverrides>>({});
    const [tab, setTab] = useState<OutputTab>("m");
    const [settingsOpen, setSettingsOpen] = useState(false);
    const [views, setViews] = useState<Record<string, ViewInfo[] | "loading" | "error">>({});
    const [preview, setPreview] = useState<PreviewState>({ loading: false, error: null, result: null, ranAt: null });
    const [options, setOptions] = useState<Map<string, ChoiceOption[]>>(new Map());
    const [optionsLoading, setOptionsLoading] = useState(false);
    const [optionsError, setOptionsError] = useState<string | null>(null);
    const optionGroupsLoaded = useRef(new Set<string>());
    const [sendTargets, setSendTargets] = useState<Array<{ id: string; name: string }>>([]);
    const [canReturn, setCanReturn] = useState(false);
    /** Whether another tool opened this one with a query: unknown until ToolBox answers. */
    const [launch, setLaunch] = useState<"pending" | "prefilled" | "normal">(hasToolbox() ? "pending" : "normal");
    /** The environment whose session is loaded; saving only starts once it is. */
    const [sessionFor, setSessionFor] = useState<string | null>(null);
    const lastEnv = useRef<string | null>(null);
    const [leftPercent, setLeftPercent] = useState(58);
    const mainRef = useRef<HTMLDivElement>(null);

    const tablesByName = useMemo(() => new Map((tables ?? []).map((t) => [t.logicalName, t])), [tables]);
    const connectionUrl = connection?.url ?? "";

    // ---------- loading ----------
    const reloadTables = useCallback(async () => {
        const generation = metadataGeneration.current;
        const current = () => generation === metadataGeneration.current;
        setTablesLoading(true);
        setTablesError(null);
        try {
            const loaded = await loadTables();
            if (current()) setTables(loaded);
        } catch (err) {
            if (current()) setTablesError(errorMessage(err));
        } finally {
            if (current()) setTablesLoading(false);
        }
    }, []);

    const resetMetadata = useCallback(() => {
        metadataGeneration.current += 1;
        setAttrCache(new Map());
        setAttrFailed({});
        setDateFormatsMissing(new Set());
        inflight.current.clear();
        setViews({});
        setOptions(new Map());
        setOptionsLoading(false);
        setOptionsError(null);
        optionGroupsLoaded.current.clear();
        setPreview({ loading: false, error: null, result: null, ranAt: null });
    }, []);

    useEffect(() => {
        resetMetadata();
        if (connection) void reloadTables();
        else setTables(null);
        // Only a different connection resets; a re-render with the same one doesn't
    }, [connection?.id, connection?.url, reloadTables, resetMetadata]);

    const ensureAttributes = useCallback(
        (entities: string[]) => {
            for (const e of entities) {
                if (!tablesByName.has(e) || attrCache.has(e) || attrFailed[e] || inflight.current.has(e)) continue;
                inflight.current.add(e);
                const generation = metadataGeneration.current;
                const current = () => generation === metadataGeneration.current;
                loadAttributes(e)
                    .then((loaded) => {
                        if (!current()) return;
                        setAttrCache((m) => new Map(m).set(e, loaded.attributes));
                        if (loaded.dateFormatsMissing) setDateFormatsMissing((s) => new Set(s).add(e));
                    })
                    .catch((err) => {
                        if (current()) setAttrFailed((f) => ({ ...f, [e]: errorMessage(err) }));
                    })
                    .finally(() => {
                        if (current()) inflight.current.delete(e);
                    });
            }
        },
        [tablesByName, attrCache, attrFailed],
    );

    // ---------- ToolBox integration ----------
    useToolboxEvents((event) => {
        if (event === "connection:updated" || event === "connection:created" || event === "connection:deleted") void refreshConnection();
        if (event === "settings:updated") void refreshTheme();
    });

    useEffect(() => {
        if (!hasToolbox()) return;
        (async () => {
            try {
                const ctx = await window.toolboxAPI.invocation?.getLaunchContext?.();
                let prefilled = true;
                if (ctx && typeof ctx.fetchXml === "string" && ctx.fetchXml.trim()) {
                    setFetchText(formatFetchXml(ctx.fetchXml));
                    setMode("fetchxml");
                } else if (ctx && typeof ctx.entityName === "string" && ctx.entityName) {
                    setSelectedTable(ctx.entityName);
                    setMode("table");
                } else {
                    prefilled = false;
                }
                setLaunch(prefilled ? "prefilled" : "normal");
                // Launched by another tool: offer to send the M back, unless the caller doesn't want a reply
                setCanReturn(callerExpectsReply(ctx));
                setSendTargets(otherTools(await window.toolboxAPI.invocation?.findToolsByCapability?.("fetchxml")));
            } catch (err) {
                console.warn("Inter-tool setup failed", err);
                setLaunch((l) => (l === "pending" ? "normal" : l));
            }
        })();
    }, []);

    // ---------- remember the work per environment ----------
    // A query handed over by another tool is a one-off: it is neither replaced by the saved session nor saved over it.
    useEffect(() => {
        if (!envKey || launch !== "normal" || !hasToolbox()) return;
        let cancelled = false;
        setSessionFor(null);
        window.toolboxAPI.settings
            .get(sessionKey(envKey))
            .then((raw: unknown) => {
                if (cancelled) return;
                // Switching to an environment with nothing saved starts clean rather than carrying the other environment's query over
                const saved = parseSession(raw) ?? (lastEnv.current && lastEnv.current !== envKey ? EMPTY_SESSION : null);
                if (saved) {
                    setMode(saved.mode);
                    setSelectedTable(saved.selectedTable);
                    setSelection(saved.selection);
                    setFetchText(saved.fetchText);
                    setViewName(saved.viewName);
                    setViewEdited(saved.viewEdited);
                    setEditorOpen(false);
                    setOverridesByCtx(saved.overridesByCtx);
                    setPreview({ loading: false, error: null, result: null, ranAt: null });
                }
                lastEnv.current = envKey;
                setSessionFor(envKey);
            })
            .catch((err: unknown) => {
                console.warn("Couldn't read the saved session", err);
                if (!cancelled) setSessionFor(envKey);
            });
        return () => {
            cancelled = true;
        };
    }, [envKey, launch]);

    const sessionValue = useMemo(
        () => serializeSession({ mode, selectedTable, selection, fetchText, viewName, viewEdited, overridesByCtx }),
        [mode, selectedTable, selection, fetchText, viewName, viewEdited, overridesByCtx],
    );
    // Saving starts once this environment's session is loaded, so it never overwrites it with another environment's work
    useDebouncedSave(envKey && sessionFor === envKey ? sessionKey(envKey) : null, sessionValue, 800);

    // A remembered or prefilled table that this environment doesn't have (any more) would leave the column list loading forever
    useEffect(() => {
        if (tables && selectedTable && !tablesByName.has(selectedTable)) setSelectedTable(null);
    }, [tables, selectedTable, tablesByName]);

    // ---------- table mode ----------
    const selectTable = useCallback((name: string) => {
        setSelectedTable(name);
        setPreview({ loading: false, error: null, result: null, ranAt: null });
    }, []);

    useEffect(() => {
        if (selectedTable) ensureAttributes([selectedTable]);
    }, [selectedTable, ensureAttributes]);

    // Default selection: primary key and primary name
    useEffect(() => {
        if (!selectedTable) return;
        const attrs = attrCache.get(selectedTable);
        if (!attrs || selection[selectedTable]) return;
        const t = tablesByName.get(selectedTable);
        const defaults = attrs.filter((a) => a.logicalName === t?.primaryIdAttribute || a.logicalName === t?.primaryNameAttribute).map((a) => a.logicalName);
        setSelection((s) => ({ ...s, [selectedTable]: defaults }));
    }, [selectedTable, attrCache, selection, tablesByName]);

    const tableAttrs = selectedTable ? attrCache.get(selectedTable) : undefined;
    const hasStatecode = !!tableAttrs?.some((a) => a.logicalName === "statecode");

    // ---------- the query ----------
    const parsed = useMemo((): { doc: Document | null; error: string | null } => {
        if (mode === "table") {
            if (!selectedTable || !tableAttrs) return { doc: null, error: null };
            const table = tablesByName.get(selectedTable);
            if (!table) return { doc: null, error: null };
            const chosen = new Set(selection[selectedTable] ?? []);
            const ordered = tableAttrs.filter((a) => chosen.has(a.logicalName)).map((a) => a.logicalName);
            return { doc: parseFetchXml(buildTableFetchXml(table, ordered, settings.activeOnly && hasStatecode)), error: null };
        }
        if (!debouncedFetch.trim()) return { doc: null, error: null };
        try {
            return { doc: parseFetchXml(debouncedFetch), error: null };
        } catch (err) {
            return { doc: null, error: err instanceof FetchXmlError ? err.message : errorMessage(err) };
        }
    }, [mode, selectedTable, tableAttrs, tablesByName, selection, settings.activeOnly, hasStatecode, debouncedFetch]);

    const entities = useMemo(() => (parsed.doc ? fetchEntities(parsed.doc) : null), [parsed.doc]);
    useEffect(() => {
        if (entities) ensureAttributes(entities.all);
    }, [entities, ensureAttributes]);

    const metadataPending = !!entities && entities.all.some((e) => tablesByName.has(e) && !attrCache.has(e) && !attrFailed[e]);

    const ctxKey = mode === "table" ? `table:${selectedTable ?? ""}` : `fetch:${entities?.root ?? ""}`;
    const overrides = useMemo(() => overridesByCtx[ctxKey] ?? {}, [overridesByCtx, ctxKey]);

    // Unticked columns, as a string so renaming a column doesn't rebuild the query
    const excludedKey = useMemo(
        () =>
            Object.entries(overrides)
                .filter(([, o]) => o.hidden)
                .map(([id]) => id)
                .sort()
                .join("\n"),
        [overrides],
    );

    const { model, modelError } = useMemo((): { model: QueryModel | null; modelError: string | null } => {
        if (!parsed.doc || !tables || metadataPending) return { model: null, modelError: null };
        try {
            const m = buildQueryModel({
                doc: parsed.doc,
                tables: tablesByName,
                attributes: attrCache,
                pageSize: settings.pageSize,
                includeRecordUrl: settings.includeRecordUrl,
                fromTablePicker: mode === "table",
                odataFilter: mode === "table" && settings.activeOnly && hasStatecode ? "statecode eq 0" : null,
                excluded: new Set(excludedKey ? excludedKey.split("\n") : []),
            });
            for (const e of entities?.all ?? []) {
                if (attrFailed[e]) m.warnings.unshift(`Couldn't load the columns of ${e}: ${attrFailed[e]}`);
            }
            const unsure = [...new Set(m.columns.filter((c) => c.attribute.kind === "datetime" && dateFormatsMissing.has(c.entity)).map((c) => c.entity))];
            for (const e of unsure) {
                m.warnings.unshift(
                    `Couldn't read the date formats of ${tablesByName.get(e)?.displayName ?? e}, so its date-only columns are treated as date and time. Reload the metadata to try again.`,
                );
            }
            return { model: m, modelError: null };
        } catch (err) {
            return { model: null, modelError: errorMessage(err) };
        }
    }, [parsed.doc, tables, metadataPending, tablesByName, attrCache, settings.pageSize, settings.includeRecordUrl, settings.activeOnly, mode, hasStatecode, entities, attrFailed, excludedKey, dateFormatsMissing]);

    const outputs = useMemo<OutputColumn[]>(() => (model ? resolveOutputColumns(model.columns, settings, overrides, !!model.recordUrlKey) : []), [model, settings, overrides]);
    const outputsByCol = useMemo(() => {
        const m = new Map<string, OutputColumn[]>();
        for (const o of outputs) m.set(o.colId, [...(m.get(o.colId) ?? []), o]);
        return m;
    }, [outputs]);

    const tooLargeToRemember = mode === "fetchxml" && fetchText.length > MAX_FETCH_LENGTH;
    const warnings = useMemo(() => {
        if (!model) return [];
        const extra = tooLargeToRemember ? [`This FetchXML is over ${MAX_FETCH_LENGTH.toLocaleString("en-US")} characters, so it isn't remembered when you close the tool. Save the M or the FetchXML to keep it.`] : [];
        return [...model.warnings, ...outputWarnings(model.columns, outputs), ...extra];
    }, [model, outputs, tooLargeToRemember]);

    const useOData = settings.queryStyle === "odata" && !!model?.odata;
    const mCode = useMemo(() => {
        if (!model) return null;
        try {
            return useOData ? generateODataQuery(model, outputs, settings, connectionUrl) : generateFetchXmlQuery(model, outputs, settings, connectionUrl);
        } catch (err) {
            return `// Couldn't generate the query: ${errorMessage(err)}`;
        }
    }, [model, outputs, settings, connectionUrl, useOData]);

    const parameterCode = useMemo(() => generateParameterQuery(connectionUrl, settings), [connectionUrl, settings]);

    // ---------- choice tables ----------
    const choiceSpecs = useMemo(() => (model ? choiceColumns(model.columns.filter((c) => !overrides[c.id]?.hidden)) : []), [model, overrides]);

    useEffect(() => {
        if (settings.choiceTables !== "static" || tab !== "choices" || choiceSpecs.length === 0) return;
        const groups = new Map<string, { entity: string; kind: AttributeInfo["kind"] }>();
        for (const s of choiceSpecs) {
            const key = `${s.entity}|${s.attribute.kind}`;
            if (!optionGroupsLoaded.current.has(key)) groups.set(key, { entity: s.entity, kind: s.attribute.kind });
        }
        if (groups.size === 0) return;
        const generation = metadataGeneration.current;
        setOptionsLoading(true);
        setOptionsError(null);
        const entries = Array.from(groups.entries());
        void Promise.allSettled(entries.map(([, g]) => loadOptions(g.entity, g.kind))).then((results) => {
            if (generation !== metadataGeneration.current) return;
            const failures: string[] = [];
            const loaded: Array<[string, ChoiceOption[]]> = [];
            results.forEach((r, i) => {
                const [key, g] = entries[i];
                if (r.status === "fulfilled") {
                    // Only a group that actually loaded counts as loaded; a failed one is tried again next time
                    optionGroupsLoaded.current.add(key);
                    for (const [attr, opts] of r.value) loaded.push([`${g.entity}.${attr}`, opts]);
                } else if (!failures.includes(errorMessage(r.reason))) {
                    failures.push(errorMessage(r.reason));
                }
            });
            setOptions((prev) => new Map([...prev, ...loaded]));
            if (failures.length > 0) setOptionsError(`Couldn't read some choice values, so live queries are shown for those: ${failures.join("; ")}`);
            setOptionsLoading(false);
        });
    }, [settings.choiceTables, tab, choiceSpecs]);

    const choiceQueries = useMemo<ChoiceQuery[]>(
        () =>
            choiceSpecs.map((spec) =>
                generateChoiceQuery(
                    spec,
                    tablesByName.get(spec.entity),
                    settings,
                    connectionUrl,
                    settings.choiceTables === "static" ? (options.get(`${spec.entity}.${spec.attribute.logicalName}`) ?? null) : null,
                ),
            ),
        [choiceSpecs, tablesByName, settings, connectionUrl, options],
    );

    // ---------- actions ----------
    const setOverride = useCallback(
        (id: string, patch: ColumnOverride) => {
            setOverridesByCtx((all) => {
                const ctx = { ...(all[ctxKey] ?? {}) };
                const next: ColumnOverride = { ...(ctx[id] ?? {}), ...patch };
                if (next.name !== undefined && next.name === "") delete next.name;
                ctx[id] = next;
                return { ...all, [ctxKey]: ctx };
            });
        },
        [ctxKey],
    );

    const toggleColumns = useCallback(
        (ids: string[], included: boolean) => {
            if (mode === "table") {
                if (!selectedTable) return;
                const names = ids.map((id) => id.replace(/^\|/, ""));
                setSelection((s) => {
                    const current = new Set(s[selectedTable] ?? []);
                    for (const n of names) {
                        if (included) current.add(n);
                        else current.delete(n);
                    }
                    return { ...s, [selectedTable]: Array.from(current) };
                });
            } else {
                setOverridesByCtx((all) => {
                    const ctx = { ...(all[ctxKey] ?? {}) };
                    for (const id of ids) ctx[id] = { ...(ctx[id] ?? {}), hidden: !included };
                    return { ...all, [ctxKey]: ctx };
                });
            }
        },
        [mode, selectedTable, ctxKey],
    );

    /**
     * Loads a query into FetchXML mode. `view` is the name of the view it comes from; "changed" keeps the
     * current view's name and marks it as edited (the query came back from another tool).
     */
    const openFetchXml = useCallback((xml: string, options: { carryOverridesFrom?: string; view?: string | "changed"; expand?: boolean } = {}) => {
        const { carryOverridesFrom } = options;
        const text = formatFetchXml(xml);
        if (options.view === "changed") setViewEdited(true);
        else {
            setViewName(options.view ?? null);
            setViewEdited(false);
        }
        setEditorOpen(!!options.expand);
        if (carryOverridesFrom) {
            try {
                const root = fetchEntities(parseFetchXml(xml)).root;
                setOverridesByCtx((all) => ({ ...all, [`fetch:${root}`]: { ...(all[carryOverridesFrom] ?? {}) } }));
            } catch {
                /* nothing to carry */
            }
        }
        setFetchText(text);
        setMode("fetchxml");
        setPreview({ loading: false, error: null, result: null, ranAt: null });
    }, []);

    const formatFetch = useCallback(() => {
        try {
            setFetchText(linesToXml(serializeLines(parseFetchXml(fetchText).documentElement)));
        } catch (err) {
            void notify("Can't format", errorMessage(err), "warning");
        }
    }, [fetchText]);

    const loadFetchFile = useCallback(async () => {
        try {
            const path = await window.toolboxAPI.fileSystem.selectPath({
                type: "file",
                title: "Open FetchXML",
                filters: [{ name: "FetchXML", extensions: ["xml", "fetchxml", "txt"] }],
            });
            if (!path) return;
            openFetchXml(await window.toolboxAPI.fileSystem.readText(path));
        } catch (err) {
            void notify("Couldn't open the file", errorMessage(err), "error");
        }
    }, [openFetchXml]);

    const sendTo = useCallback(
        async (toolId: string) => {
            try {
                const result = (await window.toolboxAPI.invocation.launchTool(toolId, { fetchXml: fetchText })) as Record<string, unknown> | null;
                const returned = result && (typeof result.fetchXml === "string" ? result.fetchXml : typeof result.fetchxml === "string" ? result.fetchxml : null);
                if (returned) {
                    openFetchXml(returned, { view: "changed" });
                    void notify("FetchXML updated", "The query came back from the other tool.", "success");
                }
            } catch (err) {
                void notify("Couldn't open the tool", errorMessage(err), "error");
            }
        },
        [fetchText, openFetchXml],
    );

    const returnToCaller = useCallback(async () => {
        if (!mCode || !model) return;
        try {
            await window.toolboxAPI.invocation.returnData({ mQuery: mCode, fetchXml: model.fetchXml });
        } catch (err) {
            void notify("Couldn't send the query back", errorMessage(err), "error");
        }
    }, [mCode, model]);

    const runPreview = useCallback(
        async (rows: number) => {
            if (!model) return;
            const generation = metadataGeneration.current;
            setPreview({ loading: true, error: null, result: null, ranAt: null });
            try {
                const records =
                    useOData && model.odata
                        ? await runODataPreview(model.table.entitySetName, odataSelect(model, outputs), model.odata.filter, rows)
                        : await runFetchPreview(model.table.entitySetName, previewFetchXml(model, rows));
                if (generation !== metadataGeneration.current) return;
                setPreview({
                    loading: false,
                    error: null,
                    result: buildPreview(model, outputs, records.slice(0, rows), settings, connectionUrl),
                    ranAt: Date.now(),
                    ranFor: mCode,
                });
            } catch (err) {
                if (generation === metadataGeneration.current) setPreview({ loading: false, error: errorMessage(err), result: null, ranAt: null });
            }
        },
        [model, outputs, settings, connectionUrl, useOData, mCode],
    );

    const openViews = useCallback(() => {
        const current = selectedTable ? views[selectedTable] : undefined;
        // Opening the menu again after a failure tries once more
        if (!selectedTable || (current && current !== "error")) return;
        const generation = metadataGeneration.current;
        const set = (value: ViewInfo[] | "error") => {
            if (generation === metadataGeneration.current) setViews((v) => ({ ...v, [selectedTable]: value }));
        };
        setViews((v) => ({ ...v, [selectedTable]: "loading" }));
        loadViews(selectedTable).then(set, () => set("error"));
    }, [selectedTable, views]);

    // ---------- splitter ----------
    const clampSplit = (percent: number) => Math.min(SPLIT_MAX, Math.max(SPLIT_MIN, percent));
    const stopDrag = useRef<(() => void) | null>(null);
    useEffect(() => () => stopDrag.current?.(), []);
    const startDrag = (e: React.PointerEvent) => {
        e.preventDefault();
        const move = (ev: PointerEvent) => {
            const rect = mainRef.current?.getBoundingClientRect();
            if (!rect) return;
            setLeftPercent(clampSplit(((ev.clientX - rect.left) / rect.width) * 100));
        };
        const up = () => {
            window.removeEventListener("pointermove", move);
            window.removeEventListener("pointerup", up);
            stopDrag.current = null;
        };
        window.addEventListener("pointermove", move);
        window.addEventListener("pointerup", up);
        stopDrag.current = up;
    };
    const splitKey = (e: React.KeyboardEvent) => {
        const next = { ArrowLeft: leftPercent - 2, ArrowRight: leftPercent + 2, Home: SPLIT_MIN, End: SPLIT_MAX }[e.key];
        if (next === undefined) return;
        e.preventDefault();
        setLeftPercent(clampSplit(next));
    };

    // ---------- grid rows ----------
    const specById = useMemo(() => new Map((model?.columns ?? []).map((c) => [c.id, c])), [model]);
    const gridRows = useMemo<GridRow[]>(() => {
        if (mode === "table") {
            if (!tableAttrs || !selectedTable) return [];
            const chosen = new Set(selection[selectedTable] ?? []);
            return tableAttrs.map((a) => {
                const id = `|${a.logicalName}`;
                return { id, attribute: a, included: chosen.has(a.logicalName), spec: chosen.has(a.logicalName) ? (specById.get(id) ?? null) : null };
            });
        }
        return (model?.columns ?? []).map((c) => ({
            id: c.id,
            attribute: c.attribute,
            included: !overrides[c.id]?.hidden,
            spec: c,
            note: [c.linkAlias ? `via ${c.linkAlias}` : null, c.attrAlias ? `as ${c.attrAlias}` : null, c.aggregate ? c.aggregate : null].filter(Boolean).join(" · ") || undefined,
        }));
    }, [mode, tableAttrs, selectedTable, selection, specById, model, overrides]);

    const queryName = (model?.table.displayName ?? "Query").replace(/[^\p{L}\p{N} _-]/gu, "").trim() || "Query";
    const currentTable = selectedTable ? tablesByName.get(selectedTable) : undefined;
    const currentViews = selectedTable ? views[selectedTable] : undefined;

    // ---------- FetchXML pane ----------
    // An empty editor has to be open to paste into, and an error has to be visible next to its cause
    const editorExpanded = editorOpen || !fetchText.trim() || !!parsed.error;
    const fetchSubject = model ? `${model.table.displayName}${model.hasLinks ? " and linked tables" : ""}` : "";
    const editorSummary = viewName ? `${viewEdited ? "Changed view" : "View"}: ${viewName}` : fetchSubject;

    // ---------- render ----------
    const viewMenu = mode === "table" && selectedTable && <ViewMenu key={selectedTable} views={currentViews} onOpen={openViews} onPick={(view) => openFetchXml(view.fetchXml, { view: view.name })} />;

    let body: React.ReactNode;
    if (!hasToolbox()) {
        body = (
            <div className={styles.center}>
                <Title3>Open this tool in Power Platform ToolBox</Title3>
                <span>It needs the ToolBox APIs to read Dataverse metadata.</span>
            </div>
        );
    } else if (connectionLoading) {
        body = <Spinner label="Checking the connection…" style={{ margin: "auto" }} />;
    } else if (!connection) {
        body = (
            <div className={styles.center}>
                <Title3>Connect to an environment</Title3>
                <span>Choose a Dataverse connection in ToolBox to list its tables.</span>
            </div>
        );
    } else {
        const emptyHint =
            mode === "table"
                ? selectedTable
                    ? metadataPending
                        ? "Loading columns…"
                        : "Pick the columns you want."
                    : "Pick a table on the left, then choose its columns. Or switch to FetchXML to paste a query with linked tables."
                : metadataPending
                  ? "Loading table metadata…"
                  : "Paste FetchXML on the left, or start from a view in Table mode.";
        body = (
            <div className={styles.main} ref={mainRef}>
                <div className={styles.left} style={{ width: `${leftPercent}%` }}>
                    {mode === "table" ? (
                        <>
                            <div className={styles.tablePane}>
                                {tablesError ? (
                                    <MessageBar intent="error" layout="multiline" style={{ margin: 8 }}>
                                        <MessageBarBody>
                                            Couldn't load tables: {tablesError}{" "}
                                            <Button size="small" icon={<ArrowClockwise16Regular />} onClick={() => void reloadTables()}>
                                                Retry
                                            </Button>
                                        </MessageBarBody>
                                    </MessageBar>
                                ) : (
                                    <TablePicker
                                        tables={tables}
                                        loading={tablesLoading}
                                        selected={selectedTable}
                                        onSelect={selectTable}
                                        onReload={() => {
                                            resetMetadata();
                                            void reloadTables();
                                        }}
                                    />
                                )}
                            </div>
                            <div className={styles.gridPane}>
                                {selectedTable && attrFailed[selectedTable] ? (
                                    <MessageBar intent="error" style={{ margin: 8 }}>
                                        <MessageBarBody>Couldn't load the columns: {attrFailed[selectedTable]}</MessageBarBody>
                                    </MessageBar>
                                ) : (
                                    <ColumnGrid
                                        title={currentTable ? currentTable.displayName : "Columns"}
                                        subtitle={currentTable && <Caption1>{currentTable.logicalName}</Caption1>}
                                        headerActions={viewMenu}
                                        rows={gridRows}
                                        loading={!!selectedTable && !tableAttrs}
                                        emptyText={selectedTable ? "This table has no readable columns." : "Pick a table to see its columns."}
                                        includeLabel="Selected"
                                        outputsByCol={outputsByCol}
                                        settings={settings}
                                        overrides={overrides}
                                        onToggle={toggleColumns}
                                        onOverride={setOverride}
                                    />
                                )}
                            </div>
                        </>
                    ) : (
                        <div className={styles.fetchPane}>
                            <div className={editorExpanded ? styles.editorPane : styles.editorPaneCollapsed}>
                                <FetchXmlEditor
                                    value={fetchText}
                                    onChange={(text) => {
                                        // Typing keeps the editor open, and the query is no longer exactly the view
                                        setFetchText(text);
                                        setEditorOpen(true);
                                        setViewEdited(true);
                                    }}
                                    expanded={editorExpanded}
                                    onToggle={() => setEditorOpen(!editorExpanded)}
                                    summary={editorSummary}
                                    error={parsed.error}
                                    onFormat={formatFetch}
                                    onLoadFile={() => void loadFetchFile()}
                                    sendTargets={sendTargets}
                                    onSendTo={(id) => void sendTo(id)}
                                />
                            </div>
                            <div style={{ flex: 1, minHeight: 0 }}>
                                <ColumnGrid
                                    title={model ? `Columns returned by ${viewName ?? fetchSubject}` : "Columns"}
                                    subtitle={model && viewName && <Caption1>{`${viewEdited ? "Changed view" : "View"} on ${fetchSubject}`}</Caption1>}
                                    rows={gridRows}
                                    loading={metadataPending}
                                    emptyText="Columns from the FetchXML appear here."
                                    includeLabel="In output"
                                    outputsByCol={outputsByCol}
                                    settings={settings}
                                    overrides={overrides}
                                    onToggle={toggleColumns}
                                    onOverride={setOverride}
                                />
                            </div>
                        </div>
                    )}
                </div>
                <div
                    className={styles.splitter}
                    onPointerDown={startDrag}
                    onKeyDown={splitKey}
                    tabIndex={0}
                    role="separator"
                    aria-orientation="vertical"
                    aria-label="Resize panes"
                    aria-valuenow={Math.round(leftPercent)}
                    aria-valuemin={SPLIT_MIN}
                    aria-valuemax={SPLIT_MAX}
                />
                <div className={styles.right}>
                    <OutputPanel
                        tab={tab}
                        onTab={setTab}
                        mode={mode}
                        model={model && (mode === "fetchxml" || (selection[selectedTable ?? ""]?.length ?? 0) > 0) ? model : null}
                        modelError={modelError}
                        warnings={warnings}
                        usesLocalTime={outputs.some((o) => o.conversion === "offset" || o.conversion === "localdate")}
                        columnCount={outputs.length + (model?.recordUrlKey ? 1 : 0)}
                        mCode={mCode}
                        settings={settings}
                        updateSettings={update}
                        parameterCode={parameterCode}
                        choiceQueries={choiceQueries}
                        choiceLoading={optionsLoading}
                        choiceError={optionsError}
                        onEditAsFetchXml={() => model && openFetchXml(model.fetchXml, { carryOverridesFrom: ctxKey, expand: true })}
                        preview={preview}
                        onRunPreview={(n) => void runPreview(n)}
                        canReturn={canReturn}
                        onReturn={() => void returnToCaller()}
                        onCopy={(t, w) => void copyText(t, w)}
                        onSave={(n, t) => void saveText(n, t)}
                        queryName={queryName}
                        emptyHint={emptyHint}
                    />
                </div>
            </div>
        );
    }

    return (
        // Layout lives on an inner div: Fluent copies the provider's className onto popup layers,
        // and a full-height flex class there would cover the page and swallow clicks.
        <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
            <div className={styles.root}>
                <header className={styles.header}>
                    <div className={styles.titleBlock}>
                        <Title3 as="h1" style={{ fontSize: 18 }}>
                            Dataverse M Builder
                        </Title3>
                        <Caption1>Dataverse to Power BI, Excel and dataflows</Caption1>
                    </div>
                    {connection && (
                        <Tooltip content={connection.url} relationship="description">
                            <Badge appearance="tint" color={ENV_COLORS[connection.environment] ?? "informative"}>
                                {connection.name} · {connection.environment}
                            </Badge>
                        </Tooltip>
                    )}
                    <span style={{ flex: 1 }} />
                    <TabList size="small" selectedValue={mode} onTabSelect={(_, d) => setMode(d.value as "table" | "fetchxml")}>
                        <Tab value="table">Table &amp; columns</Tab>
                        <Tab value="fetchxml">FetchXML</Tab>
                    </TabList>
                    <Tooltip content="Settings" relationship="label">
                        <Button appearance="subtle" icon={<Settings20Regular />} onClick={() => setSettingsOpen(true)} />
                    </Tooltip>
                </header>
                {body}
            </div>
            <SettingsDialog open={settingsOpen} onClose={() => setSettingsOpen(false)} settings={settings} update={update} reset={reset} connected={!!connection} environment={envKey} />
        </FluentProvider>
    );
}
