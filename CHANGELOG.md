# Changelog

## 1.4.0

- **Save as a Power BI project.** The Model tab saves the main query and every query that goes with it as a Power BI project (.pbip): tables with their columns and types, the URL parameter, the relationships and an empty report page. Open it in Power BI Desktop and refresh. Bridge tables and lookup keys are hidden, Record URLs are marked as links, and numbers that make sense to add up are summed by default. Exporting again gives the same files, so the project diffs cleanly in Git.
- **Save as a Power BI template.** The same model as a .pbit, which every Power BI Desktop opens without preview features, the Report Server edition included. Opening it asks for the URL parameter and loads the data into a new report.
- **Save as an Excel workbook.** _Save .xlsx_ makes a workbook with the main query, loaded to a table and refreshed when it opens. It uses Microsoft's connected-workbooks library.
- **Model tab.** It replaces the Choice tables tab and holds everything that goes with the main query: its choice tables, the new bridge and lookup tables, and the relationships between them.
- **Bridge tables for multi-select columns.** Dataverse returns the selected choices as one text ("1,3"); the bridge has one row per record and choice, so a report can filter and count by each choice. It makes its own small request (primary key and the column, under the main query's filters) and relates to the main query in both directions.
- **Lookup tables.** One query per table the lookup columns point to, with key and primary name, reading every row so records that point to an inactive row still find it. Owner combines users and teams, and Customer accounts and contacts, with a type column. Lookups to the same table share it: the first relationship is active, the others inactive for `USERELATIONSHIP`, with Owner first. A lookup to the main table itself (Parent Account) gets its own table. A switch turns lookup tables off.
- **Relationships as a TMDL script.** Paste it into Power BI Desktop's TMDL view to create all relationships at once; applying it again replaces them. The same relationships are listed for Excel's data model or for creating them by hand. Where two columns relate the same two tables, the second relationship is inactive.
- **Quick fixes.** When a choice column only returns its label, a lookup only its name, or the bridge needs the primary key, the Model tab says so and offers a button that changes the main query.
- Bridge and lookup table columns are named after their table (_Owner key_, _Account Tags bridge value_), so Power BI's relationship autodetect can't pair them with other columns. The how-to on the Model tab suggests turning autodetect off for the file anyway.
- The main query's name, which the relationships use, is remembered per table.
- The FetchXML title bar no longer repeats the view's name: the column list's title already shows it.

## 1.3.1

Fixes from a code review.

- **Table-name columns** (EntityName, such as `objecttypecode` on notes) now come out as the table's logical name. They used to default to a label, which Dataverse doesn't send for this column type, so the column was empty.
- **Date formats that can't be read are reported.** If Dataverse doesn't return a table's date formats, its date-only columns can't be told apart from date-and-time columns; the builder now says so instead of converting them silently.
- `<all-attributes />` on a table whose columns couldn't be loaded gets a clear warning that it's left out.
- **Switching environments mid-load** no longer lets results from the previous environment (columns, views, choice values, preview) land in the new one.
- **Choice values:** when one table's values fail to load, the others are still used, and the failed one is tried again next time.
- **Paging:** a paging cookie without a `pagingcookie` value is treated as no cookie, so the next page is requested by number. A fractional `count` is rounded down.
- **Unticked columns:** a linked column that a distinct query sorts on through the main table (`<order entityname="alias">`) is kept in the request.
- **Saving:** settings and your remembered work are written when the tool closes or you switch environment, instead of being lost if that happens within a second of the last change. A FetchXML too large to remember now says so.
- Smaller: the record-link app id is URL-encoded; "Record URL" is only reserved as a column name when the query has a Record URL; the table, column and view searches all match the same way (any order, ignoring accents); the pane splitter works with the keyboard.

## 1.3.0

- **The FetchXML editor collapses.** In FetchXML mode the editor folds down to its title bar once a query is loaded, which leaves the room to the column list. Click _FetchXML_ to open or close it. It opens by itself when it's empty, when you type in it, and when the XML has an error.
- **The column list names the view.** After _Start from a view_ the title reads "Columns returned by" followed by the view's name, with the table underneath; it says "Changed view" once you edit the FetchXML. The view's name is remembered with the rest of your work.

## 1.2.1

- **Unticking a column in FetchXML mode now takes it out of the query**, not only out of the result, so Dataverse no longer sends data you don't use. The request keeps the column in two cases, where removing it would change the rows: aggregate queries, and a column that a distinct query sorts on.
- Warnings about a column (not found, requested twice) disappear when you untick it. The column list still marks it.

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
