# Contributing to Dataverse M Builder

Bug reports and ideas are welcome as [GitHub issues](https://github.com/jybk/pptb-dataverse-m-builder/issues). This page covers building the tool, testing it locally, hosting it in a private marketplace and publishing a fork.

## Develop, test and load locally

```bash
npm install
npm test          # unit tests; every generated query is checked with Microsoft's M parser
npm run build     # type-check and build to dist/
npm run validate  # the same package checks the ToolBox registry runs
```

In Power Platform ToolBox: Settings → _Show Debug Menu_, then Debug → _Load Local Tool_ → choose this folder. After changes, run `npm run build` and reopen the tool tab. `npm run dev-watch` rebuilds with source maps on every save. Every push and pull request runs the tests, the build and the registry validation in GitHub Actions (`.github/workflows/ci.yml`). _Load Local Tool_ only lasts for the current ToolBox session.

## Project layout

```
src/lib/fetchxml.ts    FetchXML parsing, aliasing, normalisation and the query model
src/lib/columns.ts     output columns, modes, naming and types
src/lib/mcode.ts       M generation: FetchXML, OData, choice tables, parameter
src/lib/metadata.ts    metadata mapping
src/lib/dataverse.ts   calls to the ToolBox Dataverse API
src/lib/preview.ts     applies the M shaping to preview rows
src/lib/invocation.ts  helpers for launching and being launched by other tools
src/lib/timezone.ts    daylight saving rules (shared by the M generator and the preview) and time zone detection
src/lib/session.ts     what is remembered per environment
src/components/        React + Fluent UI components
tests/                 Vitest tests
```

## Hosting in a private marketplace

To install a modified build across your organisation without publishing it to npm:

1. Put your own name in `package.json` (`author`, `contributors`), then run `npm run pack:marketplace`. This builds the tool and writes `marketplace/pptb-dataverse-m-builder-<version>.tar.gz` and `marketplace/registry.json`.
2. Upload both files to the same folder on an HTTPS location that ToolBox can read without signing in, such as an Azure Storage static website or an internal web server. Authenticated sources aren't supported yet.
3. In ToolBox: **Settings → Marketplace sources → add a source**, type _private_, any display name, and the full URL of `registry.json`. Restart when asked.
4. Open the **Marketplace**, find the tool and install it.

To release an update, raise `version` in `package.json`, run `npm run pack:marketplace` again and upload both files. Older archives can stay in the folder.

## Publishing a fork

First put your own details in `package.json`:

- the package `name` (a scoped name such as `@your-name/tool` is recommended) and `displayName`
- `author` and `contributors`
- `configurations.repository`, `configurations.website` and `configurations.readmeUrl` (a `raw.githubusercontent.com` link to the README)

Then run `npm run validate`, `npm run build`, `npm run finalize-package` and `npm publish --access public`. Test the published package with Debug → _Install from npm_, and submit it through the ToolBox tool submission form. Read the ToolBox [Marketplace Policy](https://docs.powerplatformtoolbox.com/policies/marketplace) first: a fork needs its own name and has to credit this tool.
