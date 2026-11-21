/**
 * An Excel workbook with the main query in Power Query, loaded to a table on the first sheet and
 * refreshed when the workbook opens. Built with Microsoft's connected-workbooks library, which is
 * loaded only when a workbook is made.
 *
 * One query per workbook: that's what the library makes. The query therefore has the Dataverse URL
 * written into it rather than referring to a parameter.
 */
export async function buildWorkbook(queryName: string, code: string, columnNames: string[]): Promise<Uint8Array> {
    const { workbookManager } = await import("@microsoft/connected-workbooks");
    const blob = await workbookManager.generateSingleQueryWorkbook(
        { queryName, queryMashup: withoutNameHint(code), refreshOnOpen: true },
        // The column headers, so the table shows what's coming before the first refresh
        columnNames.length > 0 ? { data: [columnNames], config: { promoteHeaders: true, adjustColumnNames: false } } : undefined,
        { docProps: { title: queryName, createdBy: "Dataverse M Builder", description: "Power Query to Dataverse, refreshed when the workbook opens." } },
    );
    return bytesOf(blob);
}

async function bytesOf(blob: Blob): Promise<Uint8Array> {
    if (typeof blob.arrayBuffer === "function") return new Uint8Array(await blob.arrayBuffer());
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
    });
}

function withoutNameHint(code: string): string {
    return code.replace(/^\/\/ Suggested query name: .*\r?\n/, "");
}
