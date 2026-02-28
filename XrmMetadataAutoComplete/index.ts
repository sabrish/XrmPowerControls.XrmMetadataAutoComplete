import * as React from "react";
import { createRoot, type Root } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { MetadataSearchBox, ISuggestionItem, IProps } from './Components/MetadataSearchBox';

export class XrmMetadataAutoComplete implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private readonly API_VERSION = "v9.2";

	/**
	 * Snapshot of the configuration key used for the last data fetch.
	 * Format: "<metadataType>|<relatedEntity>|<filterEntity>|<outputMode>"
	 * When this changes between updateView calls, PopulateDropDown is re-triggered.
	 *
	 * Note: relatedEntity may be bound to the selectedValue output of another
	 * XrmMetadataAutoComplete control on the same form. The configKey naturally
	 * captures any change to that value, ensuring a reload when the upstream
	 * control makes a selection.
	 */
	private _loadedConfig: string = "";

	private _root: Root | null = null;
	private _divContainer: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _notifyOutputChanged: () => void;
	private _currentValue: string = "";
	private _autoCompleteValues: ISuggestionItem[] = [];

	/** Persisted loading/error state — prevents updateView re-renders from killing the spinner. */
	private _isLoading: boolean = false;
	private _errorMessage: string | null = null;

	/**
	 * True when the control requires a relatedEntity that has not been configured.
	 * Shown as a distinct hint rather than the generic "no results" message.
	 */
	private _isUnconfigured: boolean = false;

	/** AbortController for the in-flight fetch — aborted when a newer request starts. */
	private _fetchAbortController: AbortController | null = null;

	// Stable callback references — bound once so React never sees new function
	// instances on re-render.
	private readonly _onSelect = (value: string): void => {
		// The component always passes outputValue (already formatted per outputMode).
		this._currentValue = value;
		this._notifyOutputChanged();
	};

	// Only propagates an explicit clear (empty string) back to the bound field.
	private readonly _onInputChange = (value?: string): void => {
		if (value === "") {
			this._currentValue = "";
			this._notifyOutputChanged();
		}
	};

	private readonly _onRetry = (): void => {
		// Re-run the last data load without resetting _loadedConfig so that a
		// simultaneous updateView in the else-branch doesn't suppress the reload.
		this._triggerDataLoad();
	};

	constructor() { /* intentionally empty */ }

	public init(
		context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		_state: ComponentFramework.Dictionary,
		container: HTMLDivElement,
	): void {
		this._context = context;
		this._notifyOutputChanged = notifyOutputChanged;

		this._divContainer = document.createElement("div");
		container.appendChild(this._divContainer);

		this._root = createRoot(this._divContainer);

		// Render the initial empty state immediately. The first updateView call
		// (which the framework always issues after init) will trigger data loading.
		this._render();
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Always refresh context so theme changes and property updates are picked up.
		// This includes changes driven by another control's output being bound to
		// this control's relatedEntity input.
		this._context = context;

		const configKey = this._buildConfigKey();
		if (configKey !== this._loadedConfig) {
			// Configuration changed — record the new key and fetch fresh data.
			this._loadedConfig = configKey;
			this._triggerDataLoad();
		} else {
			// Config unchanged — re-render in place to propagate value or theme
			// changes (e.g. bound field reset, high-contrast toggle, disabled state).
			this._render();
		}
	}

	public getOutputs(): IOutputs {
		return { selectedValue: this._currentValue };
	}

	public destroy(): void {
		this._fetchAbortController?.abort();
		this._root?.unmount();
		this._root = null;
	}

	// ---------------------------------------------------------------------------
	// Private helpers
	// ---------------------------------------------------------------------------

	private _getRelatedEntityParam(): string | null {
		const p = this._context.parameters.relatedEntity;
		return p === undefined ? null : p.raw;
	}

	private _getFilterEntityParam(): string | null {
		const p = this._context.parameters.filterEntityFieldByEntitiesAssociatedTo;
		return p === undefined ? null : p.raw;
	}

	private _getNoSuggestionsMessage(): string {
		return this._context.parameters.noSuggestionsMessage?.raw || "No results found";
	}

	private _getSearchPlaceholder(): string {
		return this._context.parameters.searchPlaceholder?.raw || "Search...";
	}

	/**
	 * Returns "LogicalName" or "DisplayValue".
	 * LogicalName is the default and is required when this control's output is
	 * bound to another control's relatedEntity input (cascading scenario).
	 */
	private _getOutputMode(): string {
		const p = this._context.parameters.outputMode;
		return (p === undefined ? null : p.raw) || "LogicalName";
	}

	/**
	 * Builds a string key that uniquely identifies the current data-loading
	 * configuration. When this key differs from _loadedConfig, a new fetch is needed.
	 * Includes outputMode so changing the mode rebuilds the list with correct outputValues.
	 */
	private _buildConfigKey(): string {
		const metadataType = this._context.parameters.autoCompleteMetaDataType.raw;
		const relatedEntity = this._getRelatedEntityParam() ?? "";
		const filterEntity = this._getFilterEntityParam() ?? "";
		const outputMode = this._getOutputMode();
		return `${metadataType}|${relatedEntity}|${filterEntity}|${outputMode}`;
	}

	private _triggerDataLoad(): void {
		const metadataType = this._context.parameters.autoCompleteMetaDataType.raw;
		const relatedEntity = this._getRelatedEntityParam();
		const filterEntity = this._getFilterEntityParam();
		this.PopulateDropDown(metadataType, filterEntity, relatedEntity);
	}

	/** Builds props from current class state and calls _root.render(). */
	private _render(): void {
		const props: IProps = {
			value: this._context.parameters.selectedValue.raw || "",
			json: this._autoCompleteValues,
			onSelect: this._onSelect,
			onInputChange: this._onInputChange,
			onRetry: this._onRetry,
			noSuggestionMessage: this._getNoSuggestionsMessage(),
			searchTitle: this._getSearchPlaceholder(),
			isLoading: this._isLoading,
			errorMessage: this._errorMessage,
			isUnconfigured: this._isUnconfigured,
			disabled: this._context.mode.isControlDisabled,
		};
		this._root?.render(
			React.createElement(
				FluentProvider,
				{ theme: this._context.fluentDesignLanguage?.tokenTheme ?? webLightTheme },
				React.createElement(MetadataSearchBox, props),
			),
		);
	}

	// ---------------------------------------------------------------------------
	// Data loading
	// ---------------------------------------------------------------------------

	private async PopulateDropDown(
		metadataType: string,
		filterEntityFieldByEntitiesAssociatedTo: string | null,
		relatedEntityName: string | null,
	): Promise<void> {
		// Abort any in-flight request before issuing a new one.
		this._fetchAbortController?.abort();
		const controller = new AbortController();
		this._fetchAbortController = controller;

		this._isUnconfigured = false;
		this._isLoading = true;
		this._errorMessage = null;
		this._render();

		try {
			let webApiUrl: string;
			let namefield: string;
			let idField: string;

			switch (metadataType) {
				case "Entity":
					if (!filterEntityFieldByEntitiesAssociatedTo) {
						webApiUrl = this.GetEntitiesUrl();
						namefield = "LogicalName";
						idField = "MetadataId";
					} else {
						webApiUrl = this.GetOneToManyRelationshipsUrl(filterEntityFieldByEntitiesAssociatedTo);
						namefield = "ReferencingEntity";
						idField = "MetadataId";
					}
					break;
				case "Attributes":
					if (!relatedEntityName) {
						this._isLoading = false;
						this._isUnconfigured = true;
						this._autoCompleteValues = [];
						this._render();
						return;
					}
					webApiUrl = this.GetAttributesforEntityUrl(relatedEntityName);
					namefield = "LogicalName";
					idField = "MetadataId";
					break;
				case "Lookup":
					if (!relatedEntityName) {
						this._isLoading = false;
						this._isUnconfigured = true;
						this._autoCompleteValues = [];
						this._render();
						return;
					}
					webApiUrl = this.GetCustomerOrLookupAttributesforEntityUrl(relatedEntityName);
					namefield = "LogicalName";
					idField = "MetadataId";
					break;
				case "SystemViews":
					if (!relatedEntityName) {
						this._isLoading = false;
						this._isUnconfigured = true;
						this._autoCompleteValues = [];
						this._render();
						return;
					}
					webApiUrl = this.GetSavedViewsForEntityUrl(relatedEntityName);
					namefield = "name";
					idField = "savedqueryid";
					break;
				case "BusinessProcessFlows":
					if (!relatedEntityName) {
						this._isLoading = false;
						this._isUnconfigured = true;
						this._autoCompleteValues = [];
						this._render();
						return;
					}
					webApiUrl = this.GetBusinessProcessFlowsUrl(relatedEntityName);
					namefield = "name";
					idField = "workflowid";
					break;
				default:
					this._isLoading = false;
					this._errorMessage = `Unsupported metadata type: "${metadataType}"`;
					this._render();
					return;
			}

			const data = await this.getXrmMetaData(webApiUrl, controller.signal);

			// If this request was superseded by a newer one, discard the result.
			if (controller.signal.aborted) return;

			const dataJson: any[] = metadataType === "Lookup" ? (data.Attributes ?? []) : (data.value ?? []);
			const outputMode = this._getOutputMode();

			const seen = new Set<string>();
			const results: ISuggestionItem[] = [];

			for (const record of dataJson) {
				const itemKey = record[namefield] as string;
				if (!itemKey || seen.has(itemKey)) continue;
				seen.add(itemKey);

				if (metadataType === "SystemViews") {
					const searchVal = `${record[namefield]}-CRMID-${record[idField]}`;
					results.push({
						displayValue: `${record[namefield]} (${record[idField]})`,
						searchValue: searchVal,
						outputValue: searchVal, // fixed format; outputMode does not apply
					});
				} else if (metadataType === "BusinessProcessFlows") {
					results.push({
						displayValue: record[namefield],
						searchValue: record[namefield],
						outputValue: record[namefield], // outputMode does not apply
					});
				} else {
					const displayLabel: string | undefined =
						record["DisplayName"]?.UserLocalizedLabel?.Label ??
						record["DisplayName"]?.LocalizedLabels?.[0]?.Label;
					const displayValue = displayLabel ? `${displayLabel} (${itemKey})` : itemKey;
					// outputMode controls what is written to the bound field:
					//   LogicalName  → itemKey (e.g. "account") — use this when chaining controls
					//   DisplayValue → displayValue (e.g. "Account (account)")
					const outputValue = outputMode === "DisplayValue" ? displayValue : itemKey;
					results.push({ displayValue, searchValue: itemKey, outputValue });
				}
			}

			this._autoCompleteValues = results;
			this._isLoading = false;
			this._errorMessage = null;

			// If the currently saved value is no longer in the new list, clear it.
			// Compare against outputValue because that is what was written to the field.
			const currentValue = this._context.parameters.selectedValue.raw || "";
			const valueStillValid = results.some((r) => r.outputValue === currentValue);
			if (!valueStillValid && currentValue !== "") {
				this._currentValue = "";
				this._notifyOutputChanged();
			}

			this._render();

		} catch (err) {
			if (controller.signal.aborted) return;
			this._isLoading = false;
			this._errorMessage = err instanceof Error ? err.message : "Failed to load metadata";
			this._render();
		}
	}

	private async getXrmMetaData(webApiUrl: string, signal: AbortSignal): Promise<any> {
		// Prefix with the org URL so the request resolves correctly in all PCF
		// host environments, not just those where the iframe shares the org origin.
		const absoluteUrl = this._context.page.getClientUrl() + webApiUrl;
		const response = await fetch(absoluteUrl, {
			headers: {
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0",
				"Accept": "application/json",
			},
			signal,
		});
		if (!response.ok) {
			throw new Error(`Metadata fetch failed: ${response.status} ${response.statusText}`);
		}
		return response.json();
	}

	private GetEntitiesUrl(): string {
		return `/api/data/${this.API_VERSION}/EntityDefinitions?$select=LogicalName,DisplayName,MetadataId`;
	}

	private GetSavedViewsForEntityUrl(entitylogicalname: string): string {
		// querytype eq 0 = saved query (system views); $top=5000 avoids the default 50-row page limit.
		return `/api/data/${this.API_VERSION}/savedqueries?$filter=returnedtypecode eq '${entitylogicalname}' and querytype eq 0&$select=name,savedqueryid&$orderby=name asc&$top=5000`;
	}

	private GetCustomerOrLookupAttributesforEntityUrl(entitylogicalname: string): string {
		return `/api/data/${this.API_VERSION}/EntityDefinitions(LogicalName='${entitylogicalname}')?$expand=Attributes($filter=AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Lookup' or AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Customer';$select=LogicalName,DisplayName,MetadataId,AttributeType)`;
	}

	private GetAttributesforEntityUrl(entitylogicalname: string): string {
		return `/api/data/${this.API_VERSION}/EntityDefinitions(LogicalName='${entitylogicalname}')/Attributes?$select=LogicalName,DisplayName,MetadataId,AttributeType`;
	}

	private GetOneToManyRelationshipsUrl(entitylogicalname: string): string {
		return `/api/data/${this.API_VERSION}/EntityDefinitions(LogicalName='${entitylogicalname}')/OneToManyRelationships?$select=ReferencingEntity,MetadataId`;
	}

	private GetBusinessProcessFlowsUrl(entitylogicalname: string): string {
		// $top=5000 avoids the default 50-row page limit on the workflows entity endpoint.
		return `/api/data/${this.API_VERSION}/workflows?$filter=category eq 4 and primaryentity eq '${entitylogicalname}'&$select=name,workflowid&$orderby=name asc&$top=5000`;
	}
}
