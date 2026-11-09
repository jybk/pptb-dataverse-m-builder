// Helpers for the inter-tool invocation API (toolboxAPI.invocation).
import { name as PACKAGE_NAME } from "../../package.json";

/**
 * The ids ToolBox may use for this tool: the npm package name (npm, registry and local installs
 * report it as the id or as `npmPackageName`), or the bare name used by private-marketplace installs.
 */
const OWN_IDS = new Set([PACKAGE_NAME, PACKAGE_NAME.replace(/^@[^/]+\//, "")]);

export interface ToolRef {
    id: string;
    name: string;
}

/** Turns the result of `invocation.findToolsByCapability(...)` into a list of other tools, leaving out this one. */
export function otherTools(tools: unknown[] | null | undefined): ToolRef[] {
    const result: ToolRef[] = [];
    for (const item of tools ?? []) {
        if (!item || typeof item !== "object") continue;
        const t = item as {
            id?: unknown;
            name?: unknown;
            npmPackageName?: unknown;
        };
        if (typeof t.id !== "string" || !t.id) continue;
        if (OWN_IDS.has(t.id) || (typeof t.npmPackageName === "string" && OWN_IDS.has(t.npmPackageName))) continue;
        result.push({
            id: t.id,
            name: typeof t.name === "string" && t.name ? t.name : t.id,
        });
    }
    return result;
}

/**
 * Whether the tool that launched this one wants data back. `getLaunchContext()` returns the caller's
 * prefill data, with ToolBox's invocation metadata under `__pptb`; it returns null for a normal launch.
 */
export function callerExpectsReply(launchContext: Record<string, unknown> | null | undefined): boolean {
    if (!launchContext) return false;
    const meta = launchContext.__pptb as { mode?: unknown; expectsResponse?: unknown } | undefined;
    return meta?.mode !== "one-way" && meta?.expectsResponse !== false;
}
