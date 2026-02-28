# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.1.0] - 2026-02-28

### Added

- **Fluent UI v9 (`@fluentui/react-components`)** — full migration from `office-ui-fabric-react` v7.
- **React 18 `createRoot`** — replaced legacy `ReactDOM.render` with `createRoot`/`unmount`.
- **Platform libraries** — manifest now declares `<platform-library name="React" version="18.3.1" />` and `<platform-library name="Fluent" version="9.0.0" />` so the browser uses the copies already bundled by Power Platform instead of re-loading them.
- **`featureconfig.json`** — enables `pcfReactPlatformLibraries` feature flag required for platform-provided Fluent UI.
- **Dark mode / high-contrast** — `FluentProvider` receives `context.fluentDesignLanguage.tokenTheme` so all Fluent UI tokens resolve to the correct colors for the active Power Platform theme.
- **`webLightTheme` fallback** — ensures correct rendering in the standalone test harness where no platform theme is injected.
- **`outputMode` property** — new optional Enum property (`LogicalName` | `DisplayValue`) controlling what is written to the bound field on selection. `LogicalName` is the default and is required when chaining controls.
- **`searchPlaceholder` property** — configurable placeholder text for the search box.
- **`noSuggestionsMessage` property** — configurable message shown when no items match the search.
- **Cascading controls** — `relatedEntity` description updated; `configKey` change detection in `updateView` reacts automatically when a bound upstream control changes its output.
- **Loading state** — indeterminate Fluent v9 `ProgressBar` shown below the search box while a fetch is in flight.
- **Error state with Retry** — error message and Retry button rendered below the search box on fetch failure.
- **Unconfigured state** — distinct hint message when a metadata type that requires `relatedEntity` has none configured.
- **`AbortController`** — in-flight requests are cancelled when configuration changes or the control is destroyed.
- **`$top=5000`** on savedqueries and workflows endpoints — avoids the default 50-row OData page limit.
- **`$orderby=name asc`** on savedqueries and workflows endpoints.
- **Offline mock API** — `mock-api.src.js` patches `window.fetch` in the PCF test harness to return realistic mock data for all five metadata types; applied via `patch.js` postinstall hook.
- **`patch.js` postinstall script** — idempotent patcher for `node_modules/pcf-start`; injects React 18 script tag and copies mock API.
- **ESLint 9 flat config** (`eslint.config.js`) with `eslint-plugin-react-hooks`.
- **Search icon** — inline SVG magnifying glass replaces the default Combobox chevron to match Power Platform lookup field appearance.
- **Three-tier relevance ranking** in the suggestion list: exact match → starts-with → contains, searched across both logical name and display label.
- **Match highlighting** — matching text is wrapped in `<mark>` with a brand-color highlight.
- **`CLAUDE.md`** — AI assistant instructions for the project.

### Changed

- **`MetadataSearchBox.tsx`** (new file) — replaces `Autocomplete.tsx`, `ReactSearchBox.tsx`, and `Autocomplete.types.ts`; rewritten as a single functional component using Fluent v9 `Combobox`/`Option`.
- **`Autocomplete.styles.ts`** (new file, replaces `Autocomplete.style.ts`) — all styles use Fluent v9 `makeStyles` and design tokens; no hardcoded colours.
- **`index.ts`** — fully rewritten: `createRoot`, `_loadedConfig` config-key pattern for reload detection, `AbortController`, proper `async`/`await` with error handling, `_context.page.getClientUrl()` replaces deprecated `Xrm.Page.context`.
- **Manifest version** bumped to `1.1.0`.
- **`relatedEntity` description** updated to mention cascading / binding to another control's output.
- **`BusissProcessFlows` display label** corrected from `BusissProcessFlows` to `BusinessProcessFlows` (`name` attribute preserved for backward compatibility).
- **Deduplication** — replaced broken `ExistsinArray` sort-based method with a `Set<string>`.
- **Bundle size** — reduced from ~924 KiB to ~36 KiB by externalising React and Fluent UI as platform libraries.

### Fixed

- `this.firstRun == false` assignment-vs-comparison typo (was `==`, should be `=`).
- `debugger` statement left in production code in `onChangeNotify`.
- Both `@ts-ignore` suppressions removed; replaced with correct types.
- `Xrm.Page.context.getClientUrl()` replaced with `this._context.page.getClientUrl()`.
- `filterEntityFieldByEntitiesAssociatedTo` condition bug (OR of two non-null checks was always true).
- `updateView` reload condition — replaced fragile entity-name comparison with `_loadedConfig` key.

### Removed

- `office-ui-fabric-react` v7 and all `@uifabric/*` packages.
- `office-ui-fabric` v2 CSS library.
- `@fluentui/react-search` (Combobox from `@fluentui/react-components` replaces it).
- `Components/Autocomplete.tsx` — merged into `MetadataSearchBox.tsx`.
- `Components/Autocomplete.style.ts` — replaced by `Autocomplete.styles.ts` with Fluent tokens.
- `Components/Autocomplete.types.ts` — interfaces inlined or moved to `MetadataSearchBox.tsx`.
- `Components/ReactSearchBox.tsx` — merged into `MetadataSearchBox.tsx`.
- `Components/styles/colors.tsx` and `palette.tsx` — replaced by Fluent v9 design tokens.
- `ReactDOM.unmountComponentAtNode` calls — no longer needed with `createRoot`.
- `initializeIcons()` — not needed in Fluent v9 (icons are inline SVG).
- `uuidv4()` private method — was unused.

---

## [1.0.0] - prior

Initial release with `office-ui-fabric-react` v7 and React 16.
