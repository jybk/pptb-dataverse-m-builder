import { Button, Input, makeStyles, Menu, MenuGroup, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, tokens } from "@fluentui/react-components";
import { Eye16Regular, Search16Regular } from "@fluentui/react-icons";
import { useMemo, useRef, useState } from "react";
import type { ViewInfo } from "../lib/dataverse";
import { matchesSearch } from "../lib/search";

const useStyles = makeStyles({
    popover: { display: "flex", flexDirection: "column", maxHeight: "420px", minWidth: "280px", maxWidth: "420px", overflow: "hidden" },
    // The search box stays put while the list scrolls under it
    search: { flexShrink: 0, padding: tokens.spacingVerticalXS, paddingBottom: tokens.spacingVerticalS },
    list: { overflowY: "auto", minHeight: 0 },
});

/** The "Start from a view" menu: the table's system and personal views, with a search box on top. */
export function ViewMenu(props: { views: ViewInfo[] | "loading" | "error" | undefined; onOpen: () => void; onPick: (view: ViewInfo) => void }) {
    const styles = useStyles();
    const [open, setOpen] = useState(false);
    const [search, setSearch] = useState("");
    const listRef = useRef<HTMLDivElement>(null);
    const { views } = props;

    const loaded = Array.isArray(views) ? views : null;
    const matches = useMemo(() => (loaded ?? []).filter((v) => matchesSearch(v.name, search)), [loaded, search]);
    const system = matches.filter((v) => !v.personal);
    const personal = matches.filter((v) => v.personal);

    const pick = (view: ViewInfo) => {
        setOpen(false);
        setSearch("");
        props.onPick(view);
    };

    const onSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === "ArrowDown") {
            // Into the list; from there the menu's own arrow-key navigation takes over
            e.preventDefault();
            listRef.current?.querySelector<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])')?.focus();
        } else if (e.key === "Enter" && matches.length > 0) {
            // Enter takes the first match, so typing a few letters and pressing Enter is enough
            e.preventDefault();
            pick(matches[0]);
        } else if (e.key !== "Escape" && e.key !== "Tab") {
            // Keep typing (letters, space, Home/End) away from the menu's keyboard shortcuts
            e.stopPropagation();
        }
    };

    return (
        <Menu
            open={open}
            onOpenChange={(_, d) => {
                setOpen(d.open);
                if (d.open) props.onOpen();
                else setSearch("");
            }}
        >
            <MenuTrigger disableButtonEnhancement>
                <Button size="small" appearance="subtle" icon={<Eye16Regular />}>
                    Start from a view
                </Button>
            </MenuTrigger>
            <MenuPopover className={styles.popover}>
                {loaded && loaded.length > 0 && (
                    <div className={styles.search}>
                        <Input
                            size="small"
                            style={{ width: "100%" }}
                            contentBefore={<Search16Regular />}
                            placeholder="Search views"
                            aria-label="Search views"
                            value={search}
                            onChange={(_, d) => setSearch(d.value)}
                            onKeyDown={onSearchKeyDown}
                            autoFocus
                        />
                    </div>
                )}
                <MenuList ref={listRef} className={styles.list}>
                    {views === "loading" || views === undefined ? (
                        <MenuItem disabled>Loading views…</MenuItem>
                    ) : views === "error" ? (
                        <MenuItem disabled>Couldn't load views</MenuItem>
                    ) : views.length === 0 ? (
                        <MenuItem disabled>No views for this table</MenuItem>
                    ) : matches.length === 0 ? (
                        <MenuItem disabled>No views match "{search.trim()}"</MenuItem>
                    ) : (
                        <>
                            {system.length > 0 && (
                                <MenuGroup>
                                    <MenuGroupHeader>System views</MenuGroupHeader>
                                    {system.map((v) => (
                                        <MenuItem key={v.id} onClick={() => pick(v)}>
                                            {v.name}
                                            {v.isDefault ? " (default)" : ""}
                                        </MenuItem>
                                    ))}
                                </MenuGroup>
                            )}
                            {personal.length > 0 && (
                                <MenuGroup>
                                    <MenuGroupHeader>My views</MenuGroupHeader>
                                    {personal.map((v) => (
                                        <MenuItem key={v.id} onClick={() => pick(v)}>
                                            {v.name}
                                        </MenuItem>
                                    ))}
                                </MenuGroup>
                            )}
                        </>
                    )}
                </MenuList>
            </MenuPopover>
        </Menu>
    );
}
