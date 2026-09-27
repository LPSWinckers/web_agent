# Open source license notices

The web build emits `third-party-licenses.json` beside `index.html`. The Settings page loads that
static file, so the same artifact works in the hosted app and in the web client served by `npx t3`.
The manifest includes production dependencies from the web and server packages.

The generator checks installed production and optional dependencies, including workspace package
dependencies, and omits first-party `@t3tools/*` packages. During the web build it also checks
bundled module ids for imports missing from package manifests.

The build fails when a collected package has no distributable license identifier or contains no
license or notice text. Generated notices use templates from the pinned SPDX License List.
`pnpm licenses:sync` warms the ignored `.generated/` cache. Local web development skips generated
rows when the cache is missing; release builds require complete notices.

## Custom notices and package overrides

The repository-level `third-party-licenses.config.json` holds manually maintained exceptions.
Add an entry to `customNotices` for adapted icons, fonts, media, or another asset that did not come
from an npm package:

```json
{
  "name": "asset-name",
  "license": "CC-BY-4.0",
  "generatedNotices": [
    {
      "licenseId": "CC-BY-4.0",
      "preamble": ["Asset by Example Author. Changes: converted to MP3."]
    }
  ],
  "sourceUrl": "https://example.com/source",
  "bundles": ["assets", "web"]
}
```

Each `generatedNotices` item names an SPDX license template and can add `copyrights` or a short
`preamble` for attribution and provenance. Keep `noticeFile` or `noticeFiles` only when a vendored
source tree carries an intrinsic license file that should remain beside it. Paths are relative to
the config file. `bundles` controls which generated manifests include the entry and supplies the
label shown to users. Use `includeInBundles` when those differ.

Use `packageOverrides` only when an installed npm archive omits its notice or has incorrect
metadata. `version`, `license`, and `sourceUrl` are optional. Omitting `version` applies the
override to every installed version. An override can use `repositoryUrl` instead of `name` when
several packages from one monorepo share the same notice.

Fetched SPDX templates live under `.generated/` and are ignored. Do not commit or edit them;
updating dependencies or configuration is enough for the next strict build to refresh the output.
