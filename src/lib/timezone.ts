import type { BuilderSettings, DaylightSaving } from "./types";

type ZoneSettings = Pick<BuilderSettings, "utcOffsetMinutes" | "daylightSaving">;

const HOUR = 3_600_000;
const DAY = 24 * HOUR;

/** 00:00 UTC on the last Sunday of a month (month is 0-based). */
function lastSunday(year: number, month: number): number {
    const lastDay = Date.UTC(year, month + 1, 0);
    return lastDay - new Date(lastDay).getUTCDay() * DAY;
}

/** 00:00 UTC on the nth Sunday of a month (month is 0-based, n starts at 1). */
function nthSunday(year: number, month: number, n: number): number {
    const first = Date.UTC(year, month, 1);
    return first + ((7 - new Date(first).getUTCDay()) % 7) * DAY + (n - 1) * 7 * DAY;
}

/**
 * Offset from UTC, in minutes, that applies at a UTC instant. The generated M uses the same rules:
 *  - "eu": summer time from the last Sunday of March to the last Sunday of October, changing at 01:00 UTC
 *    in every European zone.
 *  - "us": daylight saving from the second Sunday of March to the first Sunday of November, changing
 *    at 02:00 local time (United States and Canada).
 */
export function localOffsetMinutes(utcMs: number, zone: ZoneSettings): number {
    const standard = zone.utcOffsetMinutes;
    if (zone.daylightSaving === "eu") {
        const year = new Date(utcMs).getUTCFullYear();
        const summer = utcMs >= lastSunday(year, 2) + HOUR && utcMs < lastSunday(year, 9) + HOUR;
        return summer ? standard + 60 : standard;
    }
    if (zone.daylightSaving === "us") {
        // Compare on the standard-time clock: it changes at 02:00 standard in March and at 01:00 standard (02:00 daylight) in November
        const standardClock = utcMs + standard * 60_000;
        const year = new Date(standardClock).getUTCFullYear();
        const daylight = standardClock >= nthSunday(year, 2, 2) + 2 * HOUR && standardClock < nthSunday(year, 10, 1) + HOUR;
        return daylight ? standard + 60 : standard;
    }
    return standard;
}

/** "UTC+01:00" */
export function offsetLabel(minutes: number): string {
    const sign = minutes < 0 ? "−" : "+";
    const abs = Math.abs(minutes);
    return `UTC${sign}${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
}

export const DAYLIGHT_SAVING_LABELS: Record<DaylightSaving, string> = {
    none: "No daylight saving",
    eu: "European summer time",
    us: "US/Canada daylight saving",
};

/**
 * The time zone of this computer, expressed as a standard offset plus one of the daylight saving rules
 * the builder knows. `offsetAt` returns the real offset (minutes east of UTC) at an instant, and is only
 * a parameter so this can be tested. Zones with other rules come back without daylight saving.
 */
export function detectTimeZone(offsetAt: (utcMs: number) => number = (ms) => -new Date(ms).getTimezoneOffset(), year = new Date().getFullYear()): ZoneSettings {
    const january = offsetAt(Date.UTC(year, 0, 15, 12));
    const july = offsetAt(Date.UTC(year, 6, 15, 12));
    if (january === july) return { utcOffsetMinutes: january, daylightSaving: "none" };
    const standard = Math.min(january, july);
    for (const rule of ["eu", "us"] as const) {
        const candidate: ZoneSettings = { utcOffsetMinutes: standard, daylightSaving: rule };
        let matches = true;
        // Every six hours through the year is enough to tell the rules apart and to reject other calendars
        for (let ms = Date.UTC(year, 0, 1); ms < Date.UTC(year + 1, 0, 1) && matches; ms += 6 * HOUR) {
            matches = localOffsetMinutes(ms, candidate) === offsetAt(ms);
        }
        if (matches) return candidate;
    }
    return { utcOffsetMinutes: standard, daylightSaving: "none" };
}
