import { describe, expect, it } from "vitest";
import { matchesSearch } from "../src/lib/search";

describe("search box matching", () => {
    it("matches everything when nothing is typed", () => {
        expect(matchesSearch("Active Accounts", "")).toBe(true);
        expect(matchesSearch("Active Accounts", "   ")).toBe(true);
    });

    it("ignores case and accents", () => {
        expect(matchesSearch("Active Accounts", "ACTIVE")).toBe(true);
        expect(matchesSearch("Opleidingsschema's één", "een")).toBe(true);
        expect(matchesSearch("Cafe klanten", "café")).toBe(true);
    });

    it("needs every typed word, in any order", () => {
        expect(matchesSearch("My Active Accounts", "accounts my")).toBe(true);
        expect(matchesSearch("My Active Accounts", "act acc")).toBe(true);
        expect(matchesSearch("My Active Accounts", "inactive accounts")).toBe(false);
    });
});
