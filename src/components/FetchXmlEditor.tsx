import {
    Button,
    makeStyles,
    Menu,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    MessageBar,
    MessageBarBody,
    Subtitle2,
    Textarea,
    tokens,
    Tooltip,
} from "@fluentui/react-components";
import { DocumentArrowUp16Regular, Open16Regular, TextAlignLeft16Regular } from "@fluentui/react-icons";
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
    bar: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, flexWrap: "wrap" },
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
    extra?: ReactNode;
}) {
    const styles = useStyles();
    return (
        <div className={styles.root}>
            <div className={styles.bar}>
                <Subtitle2>FetchXML</Subtitle2>
                <span style={{ flex: 1 }} />
                {props.extra}
                <Tooltip content="Open a .xml file" relationship="label">
                    <Button size="small" appearance="subtle" icon={<DocumentArrowUp16Regular />} onClick={props.onLoadFile}>
                        Open
                    </Button>
                </Tooltip>
                <Button size="small" appearance="subtle" icon={<TextAlignLeft16Regular />} onClick={props.onFormat} disabled={!!props.error || !props.value.trim()}>
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
            <div className={styles.area}>
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
            {props.error && (
                <MessageBar intent="error">
                    <MessageBarBody>{props.error}</MessageBarBody>
                </MessageBar>
            )}
        </div>
    );
}
