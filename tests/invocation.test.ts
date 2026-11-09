import { describe, expect, it } from "vitest";
import pkg from "../package.json";
import { callerExpectsReply, otherTools } from "../src/lib/invocation";

describe("otherTools", () => {
    const bare = pkg.name.replace(/^@[^/]+\//, "");

    it("leaves out this tool however ToolBox identifies it", () => {
        const tools = [
            { id: "fetchxml-studio", name: "FetchXML Studio" },
            { id: pkg.name, name: "Dataverse M Builder" }, // npm / Debug install
            {
                id: "registry-123",
                name: "Dataverse M Builder",
                npmPackageName: pkg.name,
            }, // official registry
            { id: bare, name: "Dataverse M Builder" }, // private marketplace
        ];
        expect(otherTools(tools)).toEqual([{ id: "fetchxml-studio", name: "FetchXML Studio" }]);
    });

    it("skips malformed entries and falls back to the id as the name", () => {
        expect(otherTools([null, 42, { name: "no id" }, { id: "" }, { id: "tool-x" }])).toEqual([{ id: "tool-x", name: "tool-x" }]);
        expect(otherTools(undefined)).toEqual([]);
    });
});

describe("callerExpectsReply", () => {
    it("is false for a normal launch", () => {
        expect(callerExpectsReply(null)).toBe(false);
        expect(callerExpectsReply(undefined)).toBe(false);
    });

    it("is true when launched by another tool without metadata, or two-way", () => {
        expect(callerExpectsReply({ fetchXml: "<fetch/>" })).toBe(true);
        expect(
            callerExpectsReply({
                fetchXml: "<fetch/>",
                __pptb: { mode: "two-way", expectsResponse: true },
            }),
        ).toBe(true);
    });

    it("is false when the caller asked for no reply", () => {
        expect(callerExpectsReply({ fetchXml: "<fetch/>", __pptb: { mode: "one-way" } })).toBe(false);
        expect(
            callerExpectsReply({
                fetchXml: "<fetch/>",
                __pptb: { expectsResponse: false },
            }),
        ).toBe(false);
    });
});
