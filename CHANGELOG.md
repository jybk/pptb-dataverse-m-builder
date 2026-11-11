# Changelog

## 1.2.0

- **Search box on the views list.** _Start from a view_ now opens with a search box above the list. Type any words from the view's name, in any order; case and accents don't matter. Enter takes the first match, the down arrow moves into the list.

## 1.1.0

### Dates and times

- **Local time with daylight saving.** "Fixed offset" is now "Local time": a standard UTC offset plus a daylight saving rule (European summer time, US/Canada daylight saving, or none). With a fixed offset, half the year was an hour off.
- **Date-only columns show the right day.** A user-local column with the _Date only_ format is stored as midnight in the user's time zone. In local time it now comes out as a `date`; in UTC the builder warns that the date can be a day off.
- The builder starts from your computer's time zone. If you already used a fixed offset, your offset is kept and daylight saving stays off until you choose a rule.
- The preview applies the same rules as the generated query.

### Working in the tool

- **Remembers your work per environment**: table, ticked columns, column names and output choices, and the FetchXML. A query handed over by another tool isn't mixed with it.
- The model-driven app for record links is now remembered per environment. Before, an app chosen in one environment was also written into queries for another, where that app doesn't exist.
- A remembered or prefilled table that the environment doesn't have no longer leaves the column list loading.
- _Start from a view_ reports when views can't be read instead of saying there are none, and tries again the next time you open it.

### Queries

- New warnings for two documented Dataverse paging limits: sorting on a linked table's column (paging may stop at 50,000 rows) and distinct queries without an `<order>`.
- Column comparisons (`valueof`) keep working when a view designer alias (`a_<guid>`) is replaced with the table name.

### Project

- Tests, build, registry validation and `npm audit` run in GitHub Actions on every push and pull request.
- `npm publish` now builds and tests first, so a stale build can't be published.

## 1.0.0 to 1.0.6

First releases: table and FetchXML modes, paging, choice tables, the URL parameter, preview, inter-tool invocation, and the marketplace listing.
