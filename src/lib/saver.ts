/**
 * Writes a value a little while after its last change, so typing doesn't cause a write per key.
 * Nothing scheduled is dropped: a value waiting under one key is written straight away when a value
 * for another key comes in (another environment), and `flush` writes whatever is waiting, for when
 * the tool is about to close.
 */
export class DebouncedSaver {
    private pending: { key: string; value: unknown } | null = null;
    private timer: ReturnType<typeof setTimeout> | null = null;

    constructor(
        private readonly write: (key: string, value: unknown) => Promise<unknown>,
        private readonly delay: number,
    ) {}

    schedule(key: string, value: unknown): void {
        if (this.pending && this.pending.key !== key) this.flush();
        this.pending = { key, value };
        if (this.timer) clearTimeout(this.timer);
        this.timer = setTimeout(() => this.flush(), this.delay);
    }

    flush(): void {
        if (this.timer) {
            clearTimeout(this.timer);
            this.timer = null;
        }
        const pending = this.pending;
        if (!pending) return;
        this.pending = null;
        this.write(pending.key, pending.value).catch((err: unknown) => console.warn(`Couldn't save ${pending.key}`, err));
    }
}
