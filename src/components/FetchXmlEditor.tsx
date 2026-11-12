import {
    Button,
    Caption1,
    makeStyles,
    Menu,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    mergeClasses,
    MessageBar,
    MessageBarBody,
    Textarea,
    tokens,
    Tooltip,
} from "@fluentui/react-components";
import { ChevronDown16Regular, ChevronRight16Regular, DocumentArrowUp16Regular, Open16Regular, TextAlignLeft16Regular } from "@fluentui/react-icons";
import type { ReactNode } from "react";

const useStyles = makeStyles({
    root: {
        display: "flex",
        flexDirection: "column",
        minHeight: 0,
        height: "100%",
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
        gap: tokens.spacingVerticalXS,
    },
    collapsed: { height: "auto" },
    bar: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, flexWrap: "wrap" },
    // The whole title is the expand/collapse button
    toggle: { minWidth: 0, paddingLeft: tokens.spacingHorizontalXS, fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
    summary: { color: tokens.colorNeutralForeground3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 },
    area: { flex: 1, minHeight: 0, display: "flex" },
    textarea: {
        flex: 1,
        "& textarea": {
            fontFamily: "'Cascadia Code', 'Cascadia Mono', Consolas, 'SF Mono', Menlo, monospace",
            fontSize: "12px",
            lineHeight: "18px",
            maxHeight: "none",
            height: "100%",
            whiteSpace: "pre",
            overflow: "auto",
        },
    },
});

export function FetchXmlEditor(props: {
    value: string;
    onChange: (value: string) => void;
    error: string | null;
    onFormat: () => void;
    onLoadFile: () => void;
    sendTargets: Array<{ id: string; name: string }>;
    onSendTo: (toolId: string) => void;
    /** Whether the editor is shown; collapsed, only the title bar remains. */
    expanded: boolean;
    onToggle: () => void;
    /** What the query is, shown next to the title (the view's name, or the table). */
    summary?: string;
    extra?: ReactNode;
}) {
    const styles = useStyles();
    return (
        <div className={mergeClasses(styles.root, !props.expanded && styles.collapsed)}>
            <div className={styles.bar}>
                <Button
                    className={styles.toggle}
                    appearance="transparent"
                    size="small"
                    icon={props.expanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
                    aria-expanded={props.expanded}
                    aria-controls="fetchxml-editor"
                    aria-label="FetchXML editor"
                    onClick={props.onToggle}
                >
                    FetchXML
                </Button>
                {props.summary && (
                    <Caption1 className={styles.summary} title={props.summary}>
                        {props.summary}
                    </Caption1>
                )}
                <span style={{ flex: 1 }} />
                {props.extra}
                <Tooltip content="Open a .xml file" relationship="label">
                    <Button size="small" appearance="subtle" icon={<DocumentArrowUp16Regular />} onClick={props.onLoadFile}>
                        Open
                    </Button>
                </Tooltip>
                <Button
                    size="small"
                    appearance="subtle"
                    icon={<TextAlignLeft16Regular />}
                    onClick={props.onFormat}
                    disabled={!!props.error || !props.value.trim()}
                >
                    Format
                </Button>
                {props.sendTargets.length > 0 && (
                    <Menu>
                        <MenuTrigger disableButtonEnhancement>
                            <Button size="small" appearance="subtle" icon={<Open16Regular />} disabled={!props.value.trim()}>
                                Edit in…
                            </Button>
                        </MenuTrigger>
                        <MenuPopover>
                            <MenuList>
                                {props.sendTargets.map((t) => (
                                    <MenuItem key={t.id} onClick={() => props.onSendTo(t.id)}>
                                        {t.name}
                                    </MenuItem>
                                ))}
                            </MenuList>
                        </MenuPopover>
                    </Menu>
                )}
            </div>
            {props.expanded && (
                <div className={styles.area} id="fetchxml-editor">
                    <Textarea
                        className={styles.textarea}
                        resize="none"
                        value={props.value}
                        onChange={(_, d) => props.onChange(d.value)}
                        placeholder={
                            'Paste FetchXML here, for example from FetchXML Studio or a view:\n\n<fetch>\n  <entity name="account">\n    <attribute name="name" />\n  </entity>\n</fetch>'
                        }
                        spellCheck={false}
                        aria-label="FetchXML"
                    />
                </div>
            )}
            {props.error && (
                <MessageBar intent="error">
                    <MessageBarBody>{props.error}</MessageBarBody>
                </MessageBar>
            )}
        </div>
    );
}
