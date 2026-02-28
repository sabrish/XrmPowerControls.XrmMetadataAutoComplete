# XrmMetadataAutoComplete

A PowerApps Component Framework (PCF) field control that provides autocomplete for
Dataverse metadata. Bind it to a `SingleLine.Text` field and let users search and
select entities, attributes, lookup fields, system views, or business process flows
without leaving the form.

![Demo](https://github.com/sabrish/XrmPowerControls.XrmMetadataAutoComplete/blob/master/XrmPowerControls.XrmMetadataAutoComplete.gif?raw=true)

---

## Features

- **Five metadata types** — Entity, Attributes, Lookup, SystemViews, BusinessProcessFlows
- **Fluent UI v9** — matches Power Platform's native look and feel; full dark mode and high-contrast support via the platform `fluentDesignLanguage` API
- **React 18** — uses platform-provided React and Fluent UI; zero-KB overhead from those libraries in the bundle
- **Cascading controls** — chain two instances (e.g. Entity picker → Attribute picker) by binding the upstream `selectedValue` field to the downstream `relatedEntity` property
- **OutputMode** — choose whether the bound field stores the logical name or the display label
- **Offline test harness** — mock API intercepts all five Dataverse metadata endpoints so the control can be tested without a live connection

---

## Properties

| Property | Type | Required | Description |
|---|---|---|---|
| `selectedValue` | `SingleLine.Text` (bound) | Yes | The field where the selected value is stored and read from. |
| `autoCompleteMetaDataType` | Enum | Yes | The type of metadata to search. See [Metadata Types](#metadata-types). |
| `relatedEntity` | `SingleLine.Text` (input) | No | Entity logical name scoping the search. Required for Attributes, Lookup, SystemViews, and BusinessProcessFlows. May be bound to another control's `selectedValue` output. |
| `filterEntityFieldByEntitiesAssociatedTo` | `SingleLine.Text` (input) | No | When `autoCompleteMetaDataType` is **Entity**, filters the list to entities that have a 1:N relationship with this entity. |
| `outputMode` | Enum | No | Controls the value written to `selectedValue`. **LogicalName** (default) stores the schema name. **DisplayValue** stores the full display label. Use `LogicalName` when chaining controls. |
| `searchPlaceholder` | `SingleLine.Text` (input) | No | Placeholder text in the search box. Defaults to `Search...`. |
| `noSuggestionsMessage` | `SingleLine.Text` (input) | No | Message shown when the search returns no matching items. Defaults to `No results found`. |

---

## Metadata Types

| Enum value | Description | `selectedValue` format |
|---|---|---|
| `Entity` | All entities (or filtered by 1:N association). | `account` |
| `Attributes` | All attributes of the Related Entity. | `name` |
| `Lookup` | Lookup, Customer, and Owner attributes of the Related Entity. | `parentaccountid` |
| `SystemViews` | System (public) saved views of the Related Entity. | `Active Accounts-CRMID-<view-guid>` |
| `BusinessProcessFlows` | Active BPFs for the Related Entity. | `Opportunity Sales Process` |

> **Note on the enum name** — The `BusissProcessFlows` enum `name` attribute contains a historical typo. It must be kept exactly as-is in the manifest to avoid breaking existing Dataverse form configurations. The display label is `BusinessProcessFlows`.

---

## Cascading Controls (chaining)

You can cascade two instances of this control on the same form so that the entity
selected in the first drives the metadata list in the second.

**Example — Attribute picker scoped to a selected entity:**

1. Add Control A bound to field `cr123_entity`, type **Entity**.
2. Add Control B bound to field `cr123_attribute`, type **Attributes**.
3. Set Control B's **Related Entity** to the field `cr123_entity` (Control A's output).
4. Set Control A's **Output Mode** to **LogicalName** (required — the downstream control needs a schema name).

When a user selects "Account" in Control A, Control B automatically reloads its list
with Account attributes.

---

## Dark Mode and High Contrast

The control reads `context.fluentDesignLanguage.tokenTheme` from the PCF context and
passes it to Fluent UI's `FluentProvider`. Power Platform automatically supplies the
correct theme object for light, dark, and high-contrast modes. No extra configuration
is needed on the control.

---

## Installation

Import the solution from the `Solution/` directory into your Dataverse environment, then
add the control to any `SingleLine.Text` field via the field's **Controls** tab in the
form editor.

---

## Development

### Prerequisites

- Node.js 18+
- [Power Platform CLI (`pac`)](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction)

### Setup

```bash
git clone https://github.com/sabrish/XrmPowerControls.XrmMetadataAutoComplete.git
cd XrmPowerControls.XrmMetadataAutoComplete
npm install        # also runs postinstall → patch.js
```

`patch.js` runs automatically after every `npm install`. It:

1. Injects a React 18 platform library script tag into `node_modules/pcf-start/index.html`.
2. Copies `mock-api.src.js` to `node_modules/pcf-start/lib/mock-api.js`.

### Running the test harness

```bash
npm start          # opens http://localhost:8181
```

The harness intercepts all `/api/data/v9.2/*` calls with realistic mock data — no live
Dataverse connection is needed.

| Metadata type | Test by setting... |
|---|---|
| Entity | Type=Entity, no Related Entity |
| Entity (filtered) | Type=Entity, Filter Entity=`account` |
| Attributes | Type=Attributes, Related Entity=`account` or `contact` |
| Lookup | Type=Lookup, Related Entity=`account` or `contact` |
| SystemViews | Type=SystemViews, Related Entity=`account` or `contact` |
| BusinessProcessFlows | Type=BusinessProcessFlows, Related Entity=`opportunity`, `lead`, `incident`, or `contact` |

To update mock data, edit `mock-api.src.js` and run `node patch.js`.

### Build

```bash
npm run build      # production build
npm run rebuild    # full clean + rebuild
```

### Linting

```bash
npx eslint XrmMetadataAutoComplete/
```

---

## Acknowledgements

- [Sriram Balaji](https://github.com/srirambalajigit) — original PCF Autocomplete reference project
- [Durgaprasad Katarti](https://github.com/durgaprasadkatari) — initial Fluent UI integration
