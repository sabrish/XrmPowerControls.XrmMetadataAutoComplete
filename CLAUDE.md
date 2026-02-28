# CLAUDE.md — XrmMetadataAutoComplete PCF Modernisation

## Project Overview

This is a PowerApps Component Framework (PCF) field control that provides autocomplete
functionality for Dataverse metadata (entities, attributes, lookups, system views, business
process flows). It is a **standard** control bound to a `SingleLine.Text` field.

The goal of this session is a **non-breaking modernisation**:
- Upgrade all deprecated dependencies
- Migrate from `office-ui-fabric-react` v7 → Fluent UI v9 (`@fluentui/react-components`),
  using the React and Fluent UI instances **provided by Power Platform** (not bundled)
- Add dark mode / high-contrast support via the PCF `fluentDesignLanguage` API
- Fix all known bugs
- Fix security and deprecation issues
- Convert class components to functional React components with hooks
- Keep the manifest property names and enum values **identical** to maintain backward
  compatibility with existing form configurations

---

## Repository Structure

```
/ (repo root)
├── package.json                          ← project-level PCF build config
├── tsconfig.json
├── pcfconfig.json
├── XrmPowerControls.XrmMetadataAutoComplete.pcfproj
├── XrmMetadataAutoComplete/
│   ├── ControlManifest.Input.xml         ← PCF manifest (properties, resources, features)
│   ├── index.ts                          ← PCF lifecycle entry point
│   └── Components/
│       ├── Autocomplete.tsx              ← Core autocomplete UI component
│       ├── Autocomplete.style.ts         ← Inline styles (to be replaced with tokens)
│       ├── Autocomplete.types.ts         ← Shared interfaces
│       └── ReactSearchBox.tsx            ← Wrapper component (to be simplified/merged)
└── Solution/                             ← Solution packaging files (do not modify)
```

---

## Build & Test Commands

```bash
# Install dependencies after package.json changes
npm install

# Run the PCF test harness (opens http://localhost:8181)
npm start

# Build for production
npm run build

# Full clean + rebuild
npm run rebuild
```

The test harness lets you test the control in a browser without deploying to Dataverse.
When testing dark mode, use the harness theme switcher or manually toggle the theme prop.

---

## Dependency Upgrade Plan

### Remove entirely
- `office-ui-fabric` (v2 CSS-only library, fully dead)
- `office-ui-fabric-react` (v7, deprecated — superseded by `@fluentui/react-components`)
- `@uifabric/styling` (used only in `Autocomplete.style.ts` for `mergeStyleSets`)

### Add
```json
"@fluentui/react-components": "^9",
"@fluentui/react-icons": "^2"
```

### Update
```json
"@types/node": "^20",
"@types/powerapps-component-framework": "^1.3",
"pcf-scripts": "^1",
"pcf-start": "^1"
```

### Do NOT add React/ReactDOM as dependencies
Power Platform provides React 18 and `@fluentui/react-components` at runtime. To use
the platform-provided instances, the manifest must declare them as external (see Manifest
section below) and the code must `import React from 'react'` normally — the bundler
will treat them as externals.

---

## Manifest Changes (`ControlManifest.Input.xml`)

### 1. Bump the control version
```xml
version="1.0.0"
```

### 2. Declare platform-provided React + Fluent UI as external
Add inside `<resources>`:
```xml
<platform-library name="React" version="18" />
<platform-library name="Fluent" version="9" />
```

### 3. Fix the BusissProcessFlows enum value typo (value attribute, NOT the name)
The `name` attribute is what gets stored in Dataverse — **do not change it** (would break
existing configurations). The `display-name-key` can be corrected:
```xml
<!-- BEFORE -->
<value name="BusissProcessFlows" display-name-key="BusissProcessFlows" ...>BusinessProcessFlows</value>

<!-- AFTER — name stays the same, only display label fixed -->
<value name="BusissProcessFlows" display-name-key="BusinessProcessFlows" description-key="Business Process Flows">BusinessProcessFlows</value>
```

> ⚠️ The switch-case in `index.ts` matches against the **inner text** (`BusinessProcessFlows`),
> not the `name` attribute. Verify this is still correct after reviewing the generated
> `ManifestTypes.d.ts`.

### 4. Enable the WebAPI feature (it's already declared — ensure it stays)
```xml
<feature-usage>
  <uses-feature name="WebAPI" required="true" />
</feature-usage>
```

---

## Bug Fixes Required

### Bug 1: Assignment vs. comparison typo (`index.ts`)
```typescript
// BEFORE (comparison, does nothing)
this.firstRun == false;

// AFTER
this.firstRun = false;
```

### Bug 2: `ExistsinArray` deduplication is broken (`index.ts`)
The method sorts a copy of the array then iterates the original indices — the comparison
is unreliable. Replace with a `Set`:

```typescript
// Replace the entire ExistsinArray method + its usage in PopulateDropDown:
const seen = new Set<string>();
for (const item of dataJson) {
  const key = item[namefield] as string;
  if (!seen.has(key)) {
    seen.add(key);
    results.push(/* ... */);
  }
}
```

### Bug 3: `debugger` statement in production (`index.ts`)
```typescript
// DELETE this line from onChangeNotify:
debugger;
```

### Bug 4: `@ts-ignore` usage
Both `@ts-ignore` comments suppress real type errors. Fix them properly:

1. The `uuidv4` function — replace the whole implementation with `crypto.randomUUID()`
   which is available in modern browsers and has correct types:
   ```typescript
   private uuidv4(): string {
     return crypto.randomUUID();
   }
   ```
   (Note: `uuidv4()` is not actually called anywhere in the current code — confirm
   this and remove the method entirely if unused.)

2. The `Xrm.Page.context` usage — see Security Fixes below.

### Bug 5: Duplicate render branch in `Autocomplete.tsx`
The `renderSearch` method has two nearly-identical JSX branches differing only by
whether `value` prop is set on `SearchBox`. Consolidate into a single render path
using a conditional `value` prop.

---

## Security & Deprecation Fixes

### Fix 1: Replace `Xrm.Page.context.getClientUrl()` (`index.ts`)

`Xrm.Page` is deprecated and removed in some environments. Use the PCF context instead:

```typescript
// BEFORE
//@ts-ignore
const serverUrl = Xrm.Page.context.getClientUrl();

// AFTER — context is already available as this._context
const serverUrl = this._context.page.getClientUrl();
```

### Fix 2: Replace raw `fetch()` with PCF WebAPI where applicable

The PCF framework provides `context.webAPI` for Dataverse OData calls, which handles
auth, base URL, and API version automatically. However, `context.webAPI` only supports
entity record CRUD — **metadata API endpoints are not supported** by `context.webAPI`.

Therefore:
- Keep raw `fetch()` for all `/api/data/vX.X/EntityDefinitions` and related metadata URLs
- **But** fix the base URL source: use `this._context.page.getClientUrl()` instead of
  the deprecated `Xrm.Page` (already covered in Fix 1 above)
- Make the API version configurable via a constant rather than hardcoded:

```typescript
private readonly API_VERSION = "v9.2";

// Usage:
const webApiUrl = `/api/data/${this.API_VERSION}/EntityDefinitions`;
```

### Fix 3: Add proper error handling to `getXrmMetaData`

```typescript
private async getXrmMetaData(webApiUrl: string): Promise<any> {
  const response = await fetch(webApiUrl, {
    headers: { 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'Accept': 'application/json' }
  });
  if (!response.ok) {
    throw new Error(`Metadata fetch failed: ${response.status} ${response.statusText}`);
  }
  return response.json();
}
```

Wrap the `PopulateDropDown` call site in a try/catch and display an error state in the
component when it fails.

### Fix 4: Remove hardcoded colours from `Autocomplete.style.ts`

The file hardcodes `#f3f2f1` (a light-mode grey) and `"black"` for focus states. These
will be replaced with Fluent UI v9 design tokens (see Dark Mode section).

---

## Dark Mode / Theming

Power Platform exposes the current theme via `context.fluentDesignLanguage` (available
since 2023 Wave 2). This object contains the Fluent UI v9 `Theme` object and a
`isDarkMode` boolean.

### Implementation approach

1. Pass `context.fluentDesignLanguage` from `index.ts` down into the React component tree.
2. Wrap the root component in `<FluentProvider theme={fluentDesignLanguage.tokenTheme}>`.
3. Replace **all** hardcoded colours in styles with Fluent UI v9 design tokens from
   `useTheme()` / `tokens` from `@fluentui/react-components`.

```typescript
// index.ts — pass theme into props
import { FluentProvider } from '@fluentui/react-components';

// In renderComponent / updateView:
const theme = context.fluentDesignLanguage?.tokenTheme;

ReactDOM.createRoot(this._divContainer).render(
  <FluentProvider theme={theme}>
    <XrmMetadataAutoCompleteApp {...props} />
  </FluentProvider>
);
```

```typescript
// In Autocomplete.tsx — use tokens instead of hardcoded colours
import { tokens, makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  listItem: {
    ':focus': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1,
      outline: 'none',
    },
  },
});
```

### Important: `FluentProvider` and `createRoot`

Because we're using platform-provided Fluent UI, only **one** `FluentProvider` should
exist per page. If Power Platform already injects one, nesting a second one is harmless
but redundant. Use it anyway for correctness in standalone/harness contexts.

---

## React Modernisation

### Replace `ReactDOM.render()` with `createRoot` (`index.ts`)

```typescript
import { createRoot, Root } from 'react-dom/client';

// Add to class:
private _root: Root | null = null;

// In init():
this._root = createRoot(this._divContainer);

// Replace all ReactDOM.render(...) calls:
this._root.render(<XrmMetadataAutoCompleteApp {...this.props} />);

// In destroy():
this._root.unmount();
this._root = null;
```

Remove all `ReactDOM.unmountComponentAtNode()` calls — they're no longer needed because
updating props via `_root.render()` re-renders in place without unmounting.

### Convert class components to functional components with hooks

Both `ReactSearchBoxV2` and `Autocomplete` should be rewritten as function components:

- `ReactSearchBoxV2` can likely be **merged into** `Autocomplete` (it's just a thin
  wrapper adding a `Fabric` container and a `Selection` object that isn't actually used
  for anything in the current UI)
- Use `useState` for `searchText`, `isSuggestionDisabled`, `value`
- Use `useCallback` for event handlers
- Use `useId()` from `@fluentui/react-components` instead of manual ID strings like
  `'SuggestionSearchBox'` and `'SuggestionContainer'`

### Replace Fluent UI v7 components with v9 equivalents

| v7 (office-ui-fabric-react) | v9 (@fluentui/react-components) |
|---|---|
| `SearchBox` | `SearchBox` from `@fluentui/react-search` or `Input` with search icon |
| `Callout` | `Popover` / `PopoverSurface` |
| `List` | Native `<ul>` with `makeStyles` or `DataGrid` |
| `FocusZone` | `useArrowNavigationGroup()` from `@fluentui/react-tabster` |
| `Fabric` | `FluentProvider` (moved to root) |
| `initializeIcons()` | Not needed — icons are SVG in v9 |
| `mergeStyleSets` | `makeStyles` from `@fluentui/react-components` |
| `IColumn`, `Selection` | Remove — unused in current rendering |
| `Stack` / `IStackTokens` | CSS gap / `makeStyles` |

---

## Component Architecture (Target State)

```
index.ts (PCF lifecycle)
└── XrmMetadataAutoCompleteApp (functional, receives theme + data)
    └── MetadataSearchBox (functional, replaces Autocomplete + ReactSearchBoxV2)
        ├── Input / SearchBox (Fluent v9)
        └── Popover > PopoverSurface > <ul> (suggestion list)
```

Files to delete after migration:
- `Components/ReactSearchBox.tsx` (merged into MetadataSearchBox)
- `Components/Autocomplete.style.ts` (replaced by makeStyles)
- `Components/Autocomplete.types.ts` (if types are inlined or re-exported from new files)

---

## Backward Compatibility Constraints

These must NOT change — they affect existing Dataverse configurations:

| Item | Constraint |
|---|---|
| Manifest `namespace` | Keep `SAB.XrmPowerControls` |
| Manifest `constructor` | Keep `XrmMetadataAutoComplete` |
| Property `name` attributes | `selectedValue`, `autoCompleteMetaDataType`, `filterEntityFieldByEntitiesAssociatedTo`, `relatedEntity` |
| Enum `name` values | `Entity`, `Attributes`, `Lookup`, `SystemViews`, `BusissProcessFlows` (typo must stay in `name`) |
| Output field | `selectedValue` — same format strings expected by consumers |
| SystemViews format | `[ViewName]-CRMID-[View GUID]` — keep exactly |

---

## Known Issues to Investigate (Not Yet Fixed)

1. **`updateView` re-render logic** — The current condition for triggering a data reload
   is fragile (comparing `relatedEntityName != this.entity`). Review whether this causes
   unnecessary API calls or misses updates when the bound field value changes without
   the related entity changing.

2. **`filterEntityFieldByEntitiesAssociatedTo` condition bug** — In `updateView` the
   condition `filterEntityFieldByEntitiesAssociatedTo != undefined || ...!= null` is
   always true (OR of two non-null checks). This was presumably meant to be AND. Review
   the original intent and fix.

3. **No loading state** — While `PopulateDropDown` awaits the fetch, the component
   renders nothing or stale data. Add a loading spinner using Fluent v9 `Spinner`.

4. **No empty-state handling for missing relatedEntity** — When `metadataType` is not
   `Entity` and `relatedEntity` is blank, the control silently renders an empty list.
   Consider showing a message.

5. **Keyboard navigation** — The `onKeyDown` handler uses `window.document.querySelector`
   with a hardcoded ID to focus the list. This is fragile. Replace with a ref and
   `useArrowNavigationGroup`.

---

## Do Not Touch

- `Solution/` directory — packaging only, no code changes needed
- `.pcfproj` file — leave as-is unless `pac` tooling requires an update
- The `selectedValue` output format for each metadata type — consumers depend on these
- The API endpoints themselves — they work correctly, only the base URL source changes

---

## Definition of Done

- [ ] `npm run build` completes with zero errors and zero `@ts-ignore` suppressions
- [ ] `npm start` harness shows the control working for all 5 metadata types
- [ ] Dark mode visually correct when harness theme is switched to dark
- [ ] High contrast mode renders legibly (test with Windows High Contrast)
- [ ] No `Xrm.Page` references remain
- [ ] No `office-ui-fabric-react` or `@uifabric` imports remain
- [ ] No `ReactDOM.render` (legacy) calls remain
- [ ] `debugger` statement removed
- [ ] All `@ts-ignore` removed
- [ ] Existing saved configurations in Dataverse still work without reconfiguring the control
