import type { AttributeInfo, ColumnSpec, QueryModel, TableInfo } from "./types";

export class FetchXmlError extends Error {}

/** Link types that only filter rows and never return columns. */
const FILTER_ONLY_LINK_TYPES = new Set(["exists", "in", "any", "not any", "all", "not all"]);

export function parseFetchXml(text: string): Document {
    const trimmed = text.trim();
    if (!trimmed) throw new FetchXmlError("Paste or type a FetchXML query.");
    const doc = new DOMParser().parseFromString(trimmed, "application/xml");
    const err = doc.getElementsByTagName("parsererror")[0];
    if (err) {
        const msg =
            (err.textContent ?? "")
                .split("\n")
                .map((s) => s.trim())
                .filter(Boolean)[0] ?? "Invalid XML";
        throw new FetchXmlError(`The FetchXML isn't valid XML: ${msg}`);
    }
    const root = doc.documentElement;
    if (!root || root.tagName !== "fetch") throw new FetchXmlError("The root element must be <fetch>.");
    const entity = childElements(root).find((c) => c.tagName === "entity");
    if (!entity || !entity.getAttribute("name")) throw new FetchXmlError('The <fetch> element needs an <entity name="..."> child.');
    return doc;
}

function childElements(el: Element): Element[] {
    return Array.from(el.childNodes).filter((n): n is Element => n.nodeType === 1);
}

/** Root table plus every linked table, so their metadata can be loaded. */
export function fetchEntities(doc: Document): { root: string; all: string[] } {
    const root = childElements(doc.documentElement).find((c) => c.tagName === "entity")!;
    const names = new Set<string>([root.getAttribute("name")!]);
    for (const link of Array.from(root.getElementsByTagName("link-entity"))) {
        const n = link.getAttribute("name");
        if (n) names.add(n);
    }
    return { root: root.getAttribute("name")!, all: Array.from(names) };
}

function escAttr(v: string): string {
    return v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function escText(v: string): string {
    return v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Serialise an element as one line per element, so it can be shown and embedded readably. */
export function serializeLines(el: Element, depth = 0, out: Array<{ depth: number; text: string }> = []): Array<{ depth: number; text: string }> {
    const attrs = Array.from(el.attributes).map((a) => `${a.name}="${escAttr(a.value)}"`);
    const open = `<${el.tagName}${attrs.length ? " " + attrs.join(" ") : ""}`;
    const kids = childElements(el);
    const text = kids.length === 0 ? (el.textContent ?? "").trim() : "";
    if (kids.length === 0 && !text) {
        out.push({ depth, text: `${open} />` });
    } else if (kids.length === 0) {
        out.push({ depth, text: `${open}>${escText(text)}</${el.tagName}>` });
    } else {
        out.push({ depth, text: `${open}>` });
        for (const k of kids) serializeLines(k, depth + 1, out);
        out.push({ depth, text: `</${el.tagName}>` });
    }
    return out;
}

export function linesToXml(lines: Array<{ depth: number; text: string }>): string {
    return lines.map((l) => "    ".repeat(l.depth) + l.text).join("\n");
}

/** FetchXML for the table picker: the chosen columns, optionally active rows only. */
export function buildTableFetchXml(table: TableInfo, attributes: string[], activeOnly: boolean): string {
    const lines = [`<fetch>`, `    <entity name="${escAttr(table.logicalName)}">`];
    for (const a of attributes) lines.push(`        <attribute name="${escAttr(a)}" />`);
    if (activeOnly) {
        lines.push(`        <filter>`, `            <condition attribute="statecode" operator="eq" value="0" />`, `        </filter>`);
    }
    lines.push(`    </entity>`, `</fetch>`);
    return lines.join("\n");
}

export interface ModelInput {
    doc: Document;
    tables: Map<string, TableInfo>;
    attributes: Map<string, AttributeInfo[]>;
    pageSize: number;
    includeRecordUrl: boolean;
    /** Set when the query came from the table picker, to allow an OData version. */
    odataFilter?: string | null;
    fromTablePicker?: boolean;
    /**
     * Ids of columns the user unticked. They stay in the model's column list (the grid shows them),
     * but they are taken out of the FetchXML where that is safe, and warnings about them are dropped.
     */
    excluded?: ReadonlySet<string>;
}

interface AttrNode {
    entity: string;
    linkAlias: string | null;
    linkLabel: string | null;
    name: string;
    alias: string | null;
    aggregate: string | null;
    dategrouping: string | null;
    /** The <attribute> element this column comes from. */
    el: Element;
}

const columnId = (n: { linkAlias: string | null; alias: string | null; name: string }) => `${n.linkAlias ?? ""}|${n.alias ?? n.name}`;

/**
 * Takes unticked columns out of the query, so Dataverse doesn't send data nobody uses.
 * Two cases keep the column in the request (it is still left out of the result):
 *  - aggregate queries, where removing a column would change the grouping;
 *  - a column a distinct query sorts on, because a distinct query can only sort on columns it returns.
 */
function removeExcluded(nodes: AttrNode[], excluded: ReadonlySet<string>, isAggregate: boolean, isDistinct: boolean) {
    if (isAggregate) return;
    for (const n of nodes) {
        if (!excluded.has(columnId(n))) continue;
        const parent = n.el.parentElement;
        if (!parent) continue;
        const sortedOn = isDistinct && childElements(parent).some((c) => c.tagName === "order" && (c.getAttribute("attribute") === n.name || (n.alias && c.getAttribute("alias") === n.alias)));
        if (!sortedOn) parent.removeChild(n.el);
    }
}

/** Web API JSON key for an attribute returned by a FetchXML query. */
export function responseKey(attr: { logicalName: string; kind: string }, linkAlias: string | null, attrAlias: string | null): string {
    if (attrAlias) return attrAlias;
    if (linkAlias) return `${linkAlias}.${attr.logicalName}`;
    if (attr.kind === "lookup") return `_${attr.logicalName}_value`;
    return attr.logicalName;
}

/** Alias the view designer generates for linked tables, e.g. a_48f82b1a8ad844bd90d915e7b3c4f263. */
const GENERATED_ALIAS = /^a_[0-9a-f]{32}$/i;

/** Attributes that only matter to the view designer, or that restate defaults. */
function stripDesignerAttributes(fetch: Element) {
    for (const a of ["version", "output-format", "userqueryid", "savedqueryid"]) fetch.removeAttribute(a);
    if (fetch.getAttribute("mapping") === "logical") fetch.removeAttribute("mapping");
    for (const a of ["distinct", "no-lock"]) if (fetch.getAttribute(a) === "false") fetch.removeAttribute(a);
    for (const link of Array.from(fetch.getElementsByTagName("link-entity"))) link.removeAttribute("visible");
}

/** Point conditions and orders that referred to a linked table by its old alias at the new one. */
function renameAliasReferences(doc: Document, from: string, to: string) {
    for (const el of Array.from(doc.getElementsByTagName("*"))) {
        if (el.getAttribute("entityname") === from) el.setAttribute("entityname", to);
        // Column comparisons name the other column as "alias.column"
        const valueOf = el.getAttribute("valueof");
        if (valueOf?.startsWith(`${from}.`)) el.setAttribute("valueof", to + valueOf.slice(from.length));
    }
}

/** Paging pitfalls Microsoft documents for FetchXML; they only matter when the query is paged. */
function pagingWarnings(entityEl: Element, isDistinct: boolean): string[] {
    const warnings: string[] = [];
    const orders = Array.from(entityEl.getElementsByTagName("order"));
    if (isDistinct && orders.length === 0) {
        warnings.push("The query is distinct but has no <order>. Dataverse needs one to page distinct results consistently: add an <order> on the columns you return.");
    }
    const onLinkedTable = orders.some((o) => o.getAttribute("entityname") || o.parentElement?.tagName === "link-entity");
    if (onLinkedTable) {
        warnings.push(
            "The query sorts on a linked table's column. Dataverse may then page without a paging cookie, which is slower and stops at 50,000 rows. Remove that <order> if the result can be larger; row order rarely matters in Power BI.",
        );
    }
    return warnings;
}

/** Put <attribute> elements first in each entity / link-entity; the order of everything else is kept. */
function hoistAttributes(el: Element) {
    const kids = childElements(el);
    const attrs = kids.filter((k) => k.tagName === "attribute");
    const first = kids.find((k) => k.tagName !== "attribute");
    if (first) for (const a of attrs) if (a.compareDocumentPosition(first) & Node.DOCUMENT_POSITION_PRECEDING) el.insertBefore(a, first);
    for (const link of kids.filter((k) => k.tagName === "link-entity")) hoistAttributes(link);
}

/**
 * Readable prefix for a linked table's columns: the lookup it's joined through ("Primary Contact",
 * "Betreft") for many-to-one links, otherwise the linked table's display name.
 */
function labelForLink(link: Element, parentEntity: string, linkedEntity: string, attributesOf: (entity: string) => AttributeInfo[], tables: Map<string, TableInfo>): string {
    const to = link.getAttribute("to");
    const from = link.getAttribute("from");
    const lookup = to ? attributesOf(parentEntity).find((a) => a.logicalName === to) : undefined;
    const linkedPk = tables.get(linkedEntity)?.primaryIdAttribute;
    if (lookup && lookup.kind === "lookup" && (!from || from === linkedPk)) return lookup.displayName;
    return tables.get(linkedEntity)?.displayName ?? linkedEntity;
}

function unknownAttribute(name: string): AttributeInfo {
    return {
        logicalName: name,
        schemaName: name,
        displayName: name,
        typeName: "Unknown",
        kind: "unknown",
        isPrimaryId: false,
        isPrimaryName: false,
        isCustom: false,
    };
}

export function buildQueryModel(input: ModelInput): QueryModel {
    const doc = input.doc.cloneNode(true) as Document;
    const fetch = doc.documentElement;
    const entityEl = childElements(fetch).find((c) => c.tagName === "entity")!;
    const rootName = entityEl.getAttribute("name")!;
    const table = input.tables.get(rootName);
    if (!table) throw new FetchXmlError(`Table "${rootName}" wasn't found in this environment, or it can't be queried through the Web API.`);

    const warnings: string[] = [];
    const isAggregate = fetch.getAttribute("aggregate") === "true";
    const isDistinct = fetch.getAttribute("distinct") === "true";
    const top = fetch.getAttribute("top");
    const singleRequest = isAggregate || !!top;

    fetch.removeAttribute("page");
    fetch.removeAttribute("paging-cookie");
    stripDesignerAttributes(fetch);
    let pageSize = Math.min(Math.max(1, Math.floor(input.pageSize) || 5000), 5000);
    if (!singleRequest) {
        const count = Number(fetch.getAttribute("count"));
        if (count > 0) {
            pageSize = Math.min(count, 5000);
            warnings.push(`The count="${count}" attribute is used as the page size; the query still returns every row.`);
        }
        fetch.removeAttribute("count");
    }
    if (!singleRequest) warnings.push(...pagingWarnings(entityEl, isDistinct));
    if (top) warnings.push(`The query has top="${top}", so it returns at most ${top} rows and isn't paged.`);
    if (isAggregate) warnings.push("Aggregate queries are sent as a single request (Dataverse limits aggregates to 50,000 source rows).");

    const attributesOf = (entity: string) => input.attributes.get(entity) ?? [];
    const nodes: AttrNode[] = [];
    const usedAliases = new Set<string>();
    for (const l of Array.from(entityEl.getElementsByTagName("link-entity"))) {
        const a = l.getAttribute("alias");
        if (a) usedAliases.add(a);
    }
    const aliasCounters = new Map<string, number>();
    let hasLinks = false;

    const visit = (el: Element, entity: string, linkAlias: string | null, linkLabel: string | null, returnsColumns: boolean) => {
        for (const child of childElements(el)) {
            if (child.tagName === "attribute") {
                if (!returnsColumns) continue;
                const name = child.getAttribute("name");
                if (!name) continue;
                nodes.push({
                    entity,
                    linkAlias,
                    linkLabel,
                    name,
                    alias: child.getAttribute("alias"),
                    aggregate: child.getAttribute("aggregate"),
                    dategrouping: child.getAttribute("dategrouping"),
                    el: child,
                });
            } else if (child.tagName === "all-attributes") {
                if (!returnsColumns) continue;
                const all = attributesOf(entity);
                for (const a of all) {
                    const attrEl = doc.createElement("attribute");
                    attrEl.setAttribute("name", a.logicalName);
                    el.insertBefore(attrEl, child);
                    nodes.push({ entity, linkAlias, linkLabel, name: a.logicalName, alias: null, aggregate: null, dategrouping: null, el: attrEl });
                }
                el.removeChild(child);
                warnings.push(
                    `<all-attributes /> on ${linkAlias ?? entity} was replaced with its ${all.length} readable columns. Pick only the columns you need for faster refreshes.`,
                );
            } else if (child.tagName === "link-entity") {
                hasLinks = true;
                const name = child.getAttribute("name") ?? "";
                const linkType = (child.getAttribute("link-type") ?? "inner").toLowerCase();
                const returns = returnsColumns && !FILTER_ONLY_LINK_TYPES.has(linkType);
                let alias = child.getAttribute("alias");
                let label = alias;
                const hasColumns = childElements(child).some((c) => c.tagName === "attribute" || c.tagName === "all-attributes");
                if ((!alias || GENERATED_ALIAS.test(alias)) && hasColumns && returns) {
                    // No alias, or a meaningless one from the view designer (a_<guid>): use the table name,
                    // so the response keys read "contact.emailaddress1", and fix any filter that used the old alias.
                    let n = aliasCounters.get(name) ?? 0;
                    let candidate: string;
                    do {
                        n += 1;
                        candidate = n === 1 ? name : `${name}${n}`;
                    } while (usedAliases.has(candidate) || candidate === rootName);
                    aliasCounters.set(name, n);
                    usedAliases.add(candidate);
                    if (alias) renameAliasReferences(doc, alias, candidate);
                    alias = candidate;
                    child.setAttribute("alias", alias);
                    label = labelForLink(child, entity, name, attributesOf, input.tables);
                }
                if (!returns && hasColumns) warnings.push(`Columns inside the "${linkType}" link to ${name} are ignored: that link type only filters rows.`);
                visit(child, name, alias, label, returns);
            }
        }
    };
    visit(entityEl, rootName, null, null, true);
    hoistAttributes(entityEl);
    const excluded = input.excluded ?? new Set<string>();
    // Before the record URL step, which puts the primary key back if it needs it
    removeExcluded(nodes, excluded, isAggregate, isDistinct);
    // Main table's columns first, then each linked table's, in document order
    const linkOrder = new Map<string | null, number>([[null, 0]]);
    for (const n of nodes) if (!linkOrder.has(n.linkAlias)) linkOrder.set(n.linkAlias, linkOrder.size);
    nodes.sort((a, b) => linkOrder.get(a.linkAlias)! - linkOrder.get(b.linkAlias)!);

    let recordUrlKey: string | null = null;
    if (input.includeRecordUrl) {
        if (isAggregate || isDistinct) {
            warnings.push("Record URLs need the primary key, which aggregate and distinct queries don't return. The Record URL column was left out.");
        } else {
            const pk = table.primaryIdAttribute;
            const hasPk = childElements(entityEl).some((c) => c.tagName === "attribute" && c.getAttribute("name") === pk && !c.getAttribute("alias"));
            if (!hasPk) {
                const pkEl = doc.createElement("attribute");
                pkEl.setAttribute("name", pk);
                entityEl.insertBefore(pkEl, entityEl.firstChild);
            }
            recordUrlKey = pk;
        }
    }

    const columns: ColumnSpec[] = [];
    const seenKeys = new Set<string>();
    for (const n of nodes) {
        const meta = attributesOf(n.entity).find((a) => a.logicalName === n.name);
        let attribute = meta ?? unknownAttribute(n.name);
        const isExcluded = excluded.has(columnId(n));
        // An unticked column isn't in the result, so there's nothing to check about it; the grid still marks it
        if (!meta && !isExcluded) {
            warnings.push(
                input.attributes.has(n.entity)
                    ? `Column "${n.name}" wasn't found on ${n.entity} (or can't be read); it's kept untyped.`
                    : `Metadata for ${n.entity} isn't loaded; "${n.name}" is kept untyped.`,
            );
        }
        if (n.aggregate) {
            const agg = n.aggregate.toLowerCase();
            if (agg === "count" || agg === "countcolumn") attribute = { ...attribute, kind: "integer", typeName: "IntegerType" };
            else if (agg === "avg") attribute = { ...attribute, kind: "double", typeName: "DoubleType" };
        }
        if (n.dategrouping) attribute = { ...attribute, kind: "integer", typeName: "IntegerType" };
        const key = responseKey(attribute, n.linkAlias, n.alias);
        if (seenKeys.has(key)) {
            if (!isExcluded) warnings.push(`"${key}" is requested more than once; the duplicate is ignored.`);
            continue;
        }
        seenKeys.add(key);
        columns.push({
            id: columnId(n),
            entity: n.entity,
            linkAlias: n.linkAlias,
            linkLabel: n.linkLabel,
            attrAlias: n.alias,
            aggregate: n.aggregate,
            dategrouping: n.dategrouping,
            attribute,
            key,
        });
    }
    if (columns.every((c) => excluded.has(c.id))) warnings.push("The query doesn't return any columns yet.");

    const fetchLines = serializeLines(fetch);
    let odata: QueryModel["odata"] = null;
    if (input.fromTablePicker && !hasLinks && !isAggregate) {
        odata = { select: columns.map((c) => c.key), filter: input.odataFilter ?? null };
    }

    return {
        table,
        fetchLines,
        fetchXml: linesToXml(fetchLines),
        columns,
        singleRequest,
        pageSize,
        isAggregate,
        hasLinks,
        recordUrlKey,
        odata,
        warnings,
    };
}

/** FetchXML limited to a handful of rows, for the in-tool preview. */
export function previewFetchXml(model: QueryModel, rows: number): string {
    const doc = new DOMParser().parseFromString(model.fetchXml, "application/xml");
    const fetch = doc.documentElement;
    if (!model.isAggregate) {
        const top = Number(fetch.getAttribute("top"));
        fetch.setAttribute("top", String(top > 0 ? Math.min(top, rows) : rows));
    }
    return new XMLSerializer().serializeToString(doc);
}
