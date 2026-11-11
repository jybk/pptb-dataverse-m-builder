/** Lower-case without accents, so "cafe" finds "Café". */
function fold(text: string): string {
    return text.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

/**
 * Whether a name matches what was typed in a search box: every typed word has to occur somewhere
 * in the name, in any order, ignoring case and accents. An empty search matches everything.
 */
export function matchesSearch(name: string, search: string): boolean {
    const words = fold(search).split(/\s+/).filter(Boolean);
    if (words.length === 0) return true;
    const folded = fold(name);
    return words.every((word) => folded.includes(word));
}
