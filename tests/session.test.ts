import { describe, expect, it } from "vitest";
import { appIdFor } from "../src/hooks/useToolbox";
import { parseSession, serializeSession, sessionKey, type Session } from "../src/lib/session";
import { DEFAULT_SETTINGS } from "../src/lib/types";

const session: Session = {
    mode: "fetchxml",
    selectedTable: "account",
    selection: { account: ["accountid", "name"] },
    fetchText: '<fetch><entity name="account" /></fetch>',
    overridesByCtx: { "table:account": { "|name": { name: "Customer" }, "|statecode": { mode: "value", hidden: true } } },
};

describe("remembered session", () => {
    it("round-trips through storage", () => {
        const stored = JSON.parse(JSON.stringify(serializeSession(session)));
        expect(parseSession(stored)).toEqual(session);
    });

    it("is kept per environment, whatever the casing of its address", () => {
        expect(sessionKey("https://Contoso.crm4.dynamics.com")).toBe(sessionKey("https://contoso.crm4.dynamics.com"));
        expect(sessionKey("https://contoso.crm4.dynamics.com")).not.toBe(sessionKey("https://contoso-dev.crm4.dynamics.com"));
    });

    it("ignores anything that isn't a session it wrote", () => {
        for (const raw of [null, undefined, "text", 42, [], {}, { v: 99, mode: "table" }]) expect(parseSession(raw)).toBeNull();
    });

    it("drops malformed parts instead of failing", () => {
        const parsed = parseSession({
            v: 1,
            mode: "nonsense",
            selectedTable: 12,
            selection: { account: ["name", 5, null], contact: "all" },
            fetchText: { not: "text" },
            overridesByCtx: { "table:account": { "|name": { name: "", mode: 3, hidden: "yes" }, "|revenue": { name: "Omzet" } }, broken: 1 },
        });
        expect(parsed).toEqual({
            mode: "table",
            selectedTable: null,
            selection: { account: ["name"] },
            fetchText: "",
            overridesByCtx: { "table:account": { "|revenue": { name: "Omzet" } } },
        });
    });

    it("doesn't store an enormous pasted query", () => {
        const stored = serializeSession({ ...session, fetchText: "x".repeat(200_001) }) as { fetchText: string };
        expect(stored.fetchText).toBe("");
    });
});

describe("record-link app per environment", () => {
    const DEV = "https://contoso-dev.crm4.dynamics.com";
    const PROD = "https://contoso.crm4.dynamics.com";

    it("uses the app chosen in that environment only", () => {
        const stored = { ...DEFAULT_SETTINGS, appIdByEnv: { [DEV]: "dev-app" } };
        expect(appIdFor(stored, DEV)).toBe("dev-app");
        expect(appIdFor(stored, PROD)).toBe("");
        expect(appIdFor(stored, null)).toBe("");
    });

    it("still honours an app saved by an earlier version", () => {
        expect(appIdFor({ ...DEFAULT_SETTINGS, appId: "old-app" }, DEV)).toBe("old-app");
    });
});
