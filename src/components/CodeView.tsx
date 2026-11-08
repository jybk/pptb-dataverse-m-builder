import { makeStyles, tokens } from "@fluentui/react-components";
import { useMemo } from "react";

const useStyles = makeStyles({
    pre: {
        margin: 0,
        padding: tokens.spacingVerticalM,
        fontFamily: "'Cascadia Code', 'Cascadia Mono', Consolas, 'SF Mono', Menlo, monospace",
        fontSize: "12px",
        lineHeight: "18px",
        whiteSpace: "pre",
        overflow: "auto",
        backgroundColor: tokens.colorNeutralBackground2,
        color: tokens.colorNeutralForeground1,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        tabSize: 4,
        flex: 1,
        minHeight: 0,
        userSelect: "text",
    },
    comment: { color: tokens.colorPaletteGreenForeground1, fontStyle: "italic" },
    string: { color: tokens.colorPaletteMarigoldForeground2 },
    keyword: { color: tokens.colorPaletteBlueForeground2, fontWeight: tokens.fontWeightSemibold },
    fn: { color: tokens.colorPaletteBerryForeground1 },
    number: { color: tokens.colorPaletteTealForeground2 },
    tag: { color: tokens.colorPaletteBlueForeground2 },
    attr: { color: tokens.colorPaletteBerryForeground1 },
});

type Token = { cls: "comment" | "string" | "keyword" | "fn" | "number" | "tag" | "attr" | null; text: string };

const M_KEYWORDS = /^(let|in|each|if|then|else|and|or|not|true|false|null|type|as|meta|try|otherwise|error|is|nullable)$/;

function tokenizeM(code: string): Token[] {
    const out: Token[] = [];
    const re = /(\/\/[^\n]*|\/\*[\s\S]*?\*\/)|("(?:[^"]|"")*")|(#"(?:[^"]|"")*")|(\b\d+(?:\.\d+)?\b)|([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)/g;
    let last = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(code))) {
        if (m.index > last) out.push({ cls: null, text: code.slice(last, m.index) });
        if (m[1]) out.push({ cls: "comment", text: m[1] });
        else if (m[2]) out.push({ cls: "string", text: m[2] });
        else if (m[3]) out.push({ cls: null, text: m[3] });
        else if (m[4]) out.push({ cls: "number", text: m[4] });
        else if (m[5]) {
            const w = m[5];
            if (M_KEYWORDS.test(w)) out.push({ cls: "keyword", text: w });
            else if (w.includes(".") && /^[A-Z]/.test(w)) out.push({ cls: "fn", text: w });
            else out.push({ cls: null, text: w });
        }
        last = re.lastIndex;
    }
    if (last < code.length) out.push({ cls: null, text: code.slice(last) });
    return out;
}

function tokenizeXml(code: string): Token[] {
    const out: Token[] = [];
    const re = /(<\/?[\w:-]+)|(\s[\w:-]+=)("[^"]*")|(\/?>)/g;
    let last = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(code))) {
        if (m.index > last) out.push({ cls: null, text: code.slice(last, m.index) });
        if (m[1]) out.push({ cls: "tag", text: m[1] });
        else if (m[2]) {
            out.push({ cls: "attr", text: m[2] });
            out.push({ cls: "string", text: m[3] });
        } else if (m[4]) out.push({ cls: "tag", text: m[4] });
        last = re.lastIndex;
    }
    if (last < code.length) out.push({ cls: null, text: code.slice(last) });
    return out;
}

export function CodeView({ code, language = "m" }: { code: string; language?: "m" | "xml" }) {
    const styles = useStyles();
    const tokens = useMemo(() => (code.length > 200_000 ? [{ cls: null, text: code } as Token] : language === "xml" ? tokenizeXml(code) : tokenizeM(code)), [code, language]);
    return (
        <pre className={styles.pre} tabIndex={0} aria-label={language === "xml" ? "FetchXML" : "Power Query M code"}>
            {tokens.map((t, i) =>
                t.cls ? (
                    <span key={i} className={styles[t.cls]}>
                        {t.text}
                    </span>
                ) : (
                    t.text
                ),
            )}
        </pre>
    );
}
