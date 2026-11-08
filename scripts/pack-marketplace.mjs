// Builds a private-marketplace package for Power Platform ToolBox:
//   marketplace/<name>-<version>.tar.gz   package.json, pptb.config.json and dist/ at the archive root
//                                         (<name> without its npm @scope)
//   marketplace/registry.json             the registry ToolBox reads (Settings > Marketplace sources)
//
// Host both files in the same HTTPS folder. ToolBox extracts the archive as-is (no npm "package/" folder),
// and resolves a relative downloadUrl against the registry's own URL.
// Usage: npm run pack:marketplace   (runs the build first)
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
if (!existsSync(join(root, "dist", "index.html"))) {
    console.error("dist/index.html is missing: run `npm run build` first.");
    process.exit(1);
}

const outDir = join(root, "marketplace");
mkdirSync(outDir, { recursive: true });
// ToolBox installs into <tools folder>/<id> and only creates the last level, so a scoped npm name
// ("@scope/name") can't be the id here. Use the bare name for the id and the archive.
const id = pkg.name.replace(/^@[^/]+\//, "");
const archive = `${id}-${pkg.version}.tar.gz`;
const files = ["package.json", "dist", ...(existsSync(join(root, "pptb.config.json")) ? ["pptb.config.json"] : [])];

// tar ships with Windows 10+, macOS and Linux; ToolBox itself uses it to extract packages.
const tar = spawnSync("tar", ["-czf", join("marketplace", archive), ...files], { cwd: root, stdio: "inherit" });
if (tar.status !== 0) {
    console.error("tar failed; is it on the PATH?");
    process.exit(tar.status ?? 1);
}

const entry = {
    id,
    name: pkg.displayName ?? id,
    description: pkg.description ?? "",
    authors: (pkg.contributors ?? []).map((c) => (typeof c === "string" ? c : c.name)).filter(Boolean),
    version: pkg.version,
    downloadUrl: archive,
    status: "active",
    license: pkg.license,
    features: pkg.features,
    repository: pkg.configurations?.repository,
    publishedAt: new Date().toISOString(),
};

// Keep other tools already listed in this registry; replace this tool's entry.
const registryPath = join(outDir, "registry.json");
let registry = { tools: [] };
if (existsSync(registryPath)) {
    try {
        registry = JSON.parse(readFileSync(registryPath, "utf8"));
    } catch {
        console.warn("Existing registry.json couldn't be read; writing a new one.");
    }
}
registry.tools = [...(registry.tools ?? []).filter((t) => t.id !== entry.id), entry];
writeFileSync(registryPath, JSON.stringify(registry, null, 4) + "\n");

console.log(`\nWrote marketplace/${archive} and marketplace/registry.json (${entry.name} ${entry.version}).`);
console.log("Upload both to the same HTTPS folder, then add the registry.json URL under ToolBox Settings > Marketplace sources.");
