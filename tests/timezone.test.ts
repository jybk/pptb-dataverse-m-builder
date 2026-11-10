import * as PQP from "@microsoft/powerquery-parser";
import { describe, expect, it } from "vitest";
import { outputWarnings, resolveOutputColumns } from "../src/lib/columns";
import { buildQueryModel, buildTableFetchXml, parseFetchXml } from "../src/lib/fetchxml";
import { withTimeZone } from "../src/hooks/useToolbox";
import { generateFetchXmlQuery } from "../src/lib/mcode";
import { buildPreview } from "../src/lib/preview";
import { detectTimeZone, localOffsetMinutes } from "../src/lib/timezone";
import { DEFAULT_SETTINGS, type BuilderSettings, type DaylightSaving } from "../src/lib/types";
import { fixture } from "./fixtures";

const HOUR = 3_600_000;

/** Real offset of an IANA zone at an instant, in minutes east of UTC, from the runtime's time zone database. */
function realOffset(timeZone: string): (utcMs: number) => number {
    const format = new Intl.DateTimeFormat("en-US", {
        timeZone,
        hourCycle: "h23",
        year: "numeric",
        month: "numeric",
        day: "numeric",
        hour: "numeric",
        minute: "numeric",
        second: "numeric",
    });
    return (utcMs) => {
        const parts = Object.fromEntries(format.formatToParts(new Date(utcMs)).map((p) => [p.type, Number(p.value)]));
        const wallClock = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
        return Math.round((wallClock - Math.floor(utcMs / 1000) * 1000) / 60_000);
    };
}

const ZONES: Array<[string, number, DaylightSaving]> = [
    ["Europe/Amsterdam", 60, "eu"],
    ["Europe/London", 0, "eu"],
    ["Europe/Helsinki", 120, "eu"],
    ["America/New_York", -300, "us"],
    ["America/Los_Angeles", -480, "us"],
    ["America/St_Johns", -210, "us"],
    ["Asia/Kolkata", 330, "none"],
    ["America/Phoenix", -420, "none"],
];

describe("daylight saving rules", () => {
    it.each(ZONES)("match the time zone database for %s, every half hour from 2019 to 2030", (zone, utcOffsetMinutes, daylightSaving) => {
        const actual = realOffset(zone);
        const mismatches: string[] = [];
        for (let ms = Date.UTC(2019, 0, 1); ms < Date.UTC(2031, 0, 1); ms += HOUR / 2) {
            const expected = actual(ms);
            const got = localOffsetMinutes(ms, { utcOffsetMinutes, daylightSaving });
            if (got !== expected && mismatches.length < 5) mismatches.push(`${new Date(ms).toISOString()}: ${got} instead of ${expected}`);
        }
        expect(mismatches).toEqual([]);
    });

    it.each(ZONES)("are detected from the computer's time zone: %s", (zone, utcOffsetMinutes, daylightSaving) => {
        expect(detectTimeZone(realOffset(zone), 2026)).toEqual({ utcOffsetMinutes, daylightSaving });
    });

    it("falls back to the standard offset for calendars it doesn't know", () => {
        // Southern hemisphere (summer time over new year) and a zone with its own dates
        expect(detectTimeZone(realOffset("Australia/Sydney"), 2026)).toEqual({ utcOffsetMinutes: 600, daylightSaving: "none" });
        expect(detectTimeZone(realOffset("Africa/Cairo"), 2026)).toEqual({ utcOffsetMinutes: 120, daylightSaving: "none" });
    });
});

describe("settings and the time zone", () => {
    const amsterdam = () => ({ utcOffsetMinutes: 60, daylightSaving: "eu" as const });

    it("starts new users from the computer's time zone", () => {
        expect(withTimeZone({}, amsterdam)).toMatchObject({ dateMode: "utc", utcOffsetMinutes: 60, daylightSaving: "eu" });
        expect(withTimeZone({ naming: "logical" }, amsterdam)).toMatchObject({ naming: "logical", daylightSaving: "eu" });
    });

    it("doesn't change the output of someone already using a fixed offset, or a saved choice", () => {
        expect(withTimeZone({ dateMode: "offset", utcOffsetMinutes: 120 }, amsterdam)).toMatchObject({ utcOffsetMinutes: 120, daylightSaving: "none" });
        expect(withTimeZone({ utcOffsetMinutes: -300, daylightSaving: "us" }, amsterdam)).toMatchObject({ utcOffsetMinutes: -300, daylightSaving: "us" });
    });

    it("survives a runtime without time zone data", () => {
        expect(
            withTimeZone({}, () => {
                throw new Error("no Intl");
            }),
        ).toEqual(DEFAULT_SETTINGS);
    });
});

describe("local time in the generated query", () => {
    function build(settings: BuilderSettings) {
        const { tables, attributes } = fixture();
        const xml = buildTableFetchXml(tables.get("account")!, ["name", "createdon", "new_duedate", "new_startdate", "new_reviewdate"], false);
        const model = buildQueryModel({ doc: parseFetchXml(xml), tables, attributes, pageSize: 5000, includeRecordUrl: false, fromTablePicker: true });
        const outputs = resolveOutputColumns(model.columns, settings, {});
        return { model, outputs, code: generateFetchXmlQuery(model, outputs, settings, "https://contoso.crm4.dynamics.com") };
    }
    async function expectValidM(code: string) {
        const result = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, code);
        if (!PQP.TaskUtils.isParseStageOk(result)) throw new Error(`Generated M doesn't parse:\n\n${code}`);
    }
    const byName = (outputs: ReturnType<typeof build>["outputs"]) => Object.fromEntries(outputs.map((o) => [o.name, [o.conversion, o.mType]]));

    it("warns that a user-local date-only column can be a day off in UTC", () => {
        const { model, outputs } = build(DEFAULT_SETTINGS);
        expect(byName(outputs)["Due Date"]).toEqual(["utc", "type datetime"]);
        const warnings = outputWarnings(model.columns, outputs);
        expect(warnings).toHaveLength(1);
        expect(warnings[0]).toContain('"Due Date" is a date-only column');
        // Real date-only and time-zone-independent columns are never shifted, so they aren't mentioned
        expect(warnings[0]).not.toContain("Start Date");
    });

    it.each<[DaylightSaving, string[]]>([
        ["none", ["SwitchZone(DateTimeZone.FromText(value), 1, 0)"]],
        ["eu", ["LastSundayOf(Date.Year(Utc), 3)", "LastSundayOf(Date.Year(Utc), 10)", "#duration(0, 2, 0, 0) else #duration(0, 1, 0, 0)"]],
        ["us", ["NthSundayOf(Year, 3, 2) + #duration(0, 2, 0, 0)", "NthSundayOf(Year, 11, 1) + #duration(0, 1, 0, 0)"]],
    ])("converts to local time with daylight saving rule %s", async (daylightSaving, expected) => {
        const { model, outputs, code } = build({ ...DEFAULT_SETTINGS, dateMode: "offset", utcOffsetMinutes: 60, daylightSaving });
        await expectValidM(code);
        for (const text of expected) expect(code).toContain(text);
        expect(byName(outputs)).toMatchObject({
            "Created On": ["offset", "type datetime"],
            "Due Date": ["localdate", "type date"], // local midnight becomes the date users see
            "Start Date": ["date", "type date"],
            "Review Date": ["wallclock", "type datetime"],
        });
        expect(code).toContain("ToLocalDate = (value) => Date.From(ToLocalDateTime(value))");
        expect(code).toMatch(/"Due Date",\s+type date,\s+ToLocalDate\}/);
        expect(outputWarnings(model.columns, outputs)).toEqual([]);
    });

    it("leaves the local-time helpers out when no column needs them", async () => {
        const { code } = build({ ...DEFAULT_SETTINGS, dateMode: "formatted" });
        await expectValidM(code);
        expect(code).not.toContain("ToLocalDateTime");
        expect(code).not.toContain("LastSundayOf");
    });

    it("previews the same values the query will produce", () => {
        const settings: BuilderSettings = { ...DEFAULT_SETTINGS, dateMode: "offset", utcOffsetMinutes: 60, daylightSaving: "eu" };
        const { model, outputs } = build(settings);
        const records = [
            // Entered in Amsterdam as 1 July and 1 December: midnight local time
            { name: "Summer", createdon: "2026-06-30T22:30:00Z", new_duedate: "2026-06-30T22:00:00Z" },
            { name: "Winter", createdon: "2026-11-30T23:30:00Z", new_duedate: "2026-11-30T23:00:00Z" },
            // The hour the clocks go forward: 00:59:59 UTC is still winter time, 01:00 UTC is summer time
            { name: "Change", createdon: "2026-03-29T00:59:59Z", new_duedate: "2026-03-29T01:00:00Z" },
        ];
        const preview = buildPreview(model, outputs, records, settings, "https://contoso.crm4.dynamics.com");
        const col = (name: string) => preview.rows.map((r) => r[preview.columns.indexOf(name)]);
        expect(col("Created On")).toEqual(["2026-07-01 00:30:00", "2026-12-01 00:30:00", "2026-03-29 01:59:59"]);
        expect(col("Due Date")).toEqual(["2026-07-01", "2026-12-01", "2026-03-29"]);

        // With a fixed offset the summer date is a day early, which is what the daylight saving rule fixes
        const fixed = { ...settings, daylightSaving: "none" as const };
        const fixedPreview = buildPreview(model, resolveOutputColumns(model.columns, fixed, {}), records, fixed, "https://contoso.crm4.dynamics.com");
        expect(fixedPreview.rows[0][fixedPreview.columns.indexOf("Due Date")]).toBe("2026-06-30");
    });
});
