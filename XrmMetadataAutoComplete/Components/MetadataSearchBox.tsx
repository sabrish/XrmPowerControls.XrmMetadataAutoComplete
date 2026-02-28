import * as React from 'react';
import { useCallback, useMemo, useState } from 'react';
import {
  Combobox,
  Option,
  ProgressBar,
  type ComboboxProps,
} from '@fluentui/react-components';
import { useAutocompleteStyles } from './Autocomplete.styles';

// Inline Fluent "Search" icon SVG — avoids bundling @fluentui/react-icons.
// Matches the magnifying glass Power Platform shows on lookup fields.
const SearchIcon = (): React.ReactElement => (
  <svg
    aria-hidden="true"
    width="16"
    height="16"
    viewBox="0 0 20 20"
    xmlns="http://www.w3.org/2000/svg"
    style={{ display: 'block' }}
  >
    <path
      d="M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11Zm4.89 9.597 3.507 3.506-.707.707-3.506-3.507a6.5 6.5 0 1 1 .707-.706Z"
      fill="currentColor"
    />
  </svg>
);

export interface ISuggestionItem {
  /** Human-readable label shown in the dropdown list. */
  displayValue: string;
  /** Logical name / key used for filtering. Always the schema name. */
  searchValue: string;
  /**
   * Value written to the bound field on selection.
   * Equals searchValue when outputMode is "LogicalName" (default).
   * Equals displayValue when outputMode is "DisplayValue".
   * For SystemViews and BusinessProcessFlows the format is fixed regardless of outputMode.
   *
   * Important: when this control's output feeds another control's relatedEntity,
   * outputMode MUST be "LogicalName" so the downstream control receives a valid schema name.
   */
  outputValue: string;
}

export interface IProps {
  value: string;
  json: ISuggestionItem[];
  /** Called when the user commits a selection from the dropdown. */
  onSelect: (value: string) => void;
  /** Called on every keystroke; also called with '' when the input is cleared. */
  onInputChange: (value?: string) => void;
  /** Called when the user clicks the Retry button in the error state. */
  onRetry?: () => void;
  noSuggestionMessage: string;
  searchTitle: string;
  isLoading?: boolean;
  errorMessage?: string | null;
  /** True when the control needs a relatedEntity that has not been configured. */
  isUnconfigured?: boolean;
  /** Mirrors context.mode.isControlDisabled — disables input on locked/read-only forms. */
  disabled?: boolean;
}

// ---------------------------------------------------------------------------
// Helper: highlight query matches inside displayValue text
// Matches all occurrences (case-insensitive) and wraps them in <mark>.
// ---------------------------------------------------------------------------
function getHighlightedText(
  text: string,
  query: string,
  highlightClass: string,
): React.ReactNode {
  const trimmed = query.trim();
  if (!trimmed) return text;

  const lowerText = text.toLowerCase();
  const lowerQuery = trimmed.toLowerCase();
  const parts: React.ReactNode[] = [];
  let lastIdx = 0;
  let matchIdx = lowerText.indexOf(lowerQuery);
  let key = 0;

  while (matchIdx !== -1) {
    if (matchIdx > lastIdx) {
      parts.push(text.slice(lastIdx, matchIdx));
    }
    parts.push(
      <mark key={key++} className={highlightClass}>
        {text.slice(matchIdx, matchIdx + trimmed.length)}
      </mark>
    );
    lastIdx = matchIdx + trimmed.length;
    matchIdx = lowerText.indexOf(lowerQuery, lastIdx);
  }

  if (lastIdx < text.length) {
    parts.push(text.slice(lastIdx));
  }

  return parts.length > 1 ? <>{parts}</> : (parts[0] ?? text);
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------
export function MetadataSearchBox({
  value: propValue,
  json,
  onSelect,
  onInputChange,
  onRetry,
  noSuggestionMessage,
  searchTitle,
  isLoading = false,
  errorMessage = null,
  isUnconfigured = false,
  disabled = false,
}: IProps): React.ReactElement {
  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------
  const [inputValue, setInputValue] = useState(propValue);
  const [prevPropValue, setPrevPropValue] = useState(propValue);

  // Derived-state sync: when the PCF framework resets propValue externally (e.g.
  // the bound field is cleared, the entity changes, or an upstream control in a
  // cascading chain makes a new selection) update inputValue during the same
  // render pass — avoids the extra paint that useEffect would cause.
  if (prevPropValue !== propValue) {
    setPrevPropValue(propValue);
    setInputValue(propValue);
  }

  // ---------------------------------------------------------------------------
  // Styles
  // ---------------------------------------------------------------------------
  const styles = useAutocompleteStyles();

  // ---------------------------------------------------------------------------
  // Derived data
  // ---------------------------------------------------------------------------

  // Sort once when the full list changes, not on every keystroke.
  const sortedJson = useMemo(
    () => [...json].sort((a, b) => a.searchValue.localeCompare(b.searchValue)),
    [json]
  );

  // Three-tier relevance ranking: exact match → starts-with → contains.
  // Both searchValue (logical name) and displayValue (human label) are matched
  // so users can type either the schema name or the readable label.
  const { filteredItems, hasNoResults } = useMemo(() => {
    const lower = inputValue.toLowerCase();

    if (lower === '') {
      return { filteredItems: sortedJson, hasNoResults: sortedJson.length === 0 };
    }

    const exact: ISuggestionItem[] = [];
    const startsWith: ISuggestionItem[] = [];
    const contains: ISuggestionItem[] = [];

    for (const item of sortedJson) {
      const sv = item.searchValue.toLowerCase();
      const dv = item.displayValue.toLowerCase();
      if (sv === lower || dv === lower) {
        exact.push(item);
      } else if (sv.startsWith(lower) || dv.startsWith(lower)) {
        startsWith.push(item);
      } else if (sv.includes(lower) || dv.includes(lower)) {
        contains.push(item);
      }
    }

    const filtered = [...exact, ...startsWith, ...contains];
    return { filteredItems: filtered, hasNoResults: filtered.length === 0 };
  }, [sortedJson, inputValue]);

  // ---------------------------------------------------------------------------
  // Handlers
  // ---------------------------------------------------------------------------

  // Called on every keystroke in the Combobox input.
  const handleChange: NonNullable<ComboboxProps['onChange']> = useCallback(
    (ev) => {
      const text = ev.target.value;
      setInputValue(text);
      // _onInputChange in index.ts only acts when value === '' (to clear the field).
      // Passing the text for every keystroke is harmless for non-empty values.
      onInputChange(text);
    },
    [onInputChange]
  );

  // Called when the user selects an option from the dropdown or clears via X button.
  const handleOptionSelect: NonNullable<ComboboxProps['onOptionSelect']> = useCallback(
    (_ev, data) => {
      if (data.optionValue === undefined) {
        // X (clear) button was clicked.
        setInputValue('');
        onSelect('');
        return;
      }
      // text prop on <Option> controls what's shown in the input after selection.
      const text = data.optionText ?? data.optionValue;
      setInputValue(text);
      onSelect(data.optionValue);
    },
    [onSelect]
  );

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  return (
    <div className={styles.container}>
      {/*
        Combobox with freeform allows typing values that aren't in the list, which
        is needed so users can filter by typing partial text before selecting.
        clearable shows an X dismiss button when a value is present.
        selectedOptions tracks which option is highlighted in the dropdown — it mirrors
        propValue (the last committed PCF field value) rather than inputValue (current
        typed text) so the checkmark is stable during typing.
      */}
      <Combobox
        className={styles.combobox}
        freeform
        clearable
        placeholder={searchTitle}
        value={inputValue}
        selectedOptions={propValue ? [propValue] : []}
        onChange={handleChange}
        onOptionSelect={handleOptionSelect}
        disabled={disabled}
        aria-label={searchTitle}
        expandIcon={{ children: <SearchIcon /> }}
      >
        {isLoading ? (
          <Option key="__loading__" disabled value="">
            Loading\u2026
          </Option>
        ) : errorMessage ? (
          <Option key="__error__" disabled value="">
            Failed to load options.
          </Option>
        ) : isUnconfigured ? (
          <Option key="__unconfigured__" disabled value="">
            Set a Related Entity property to load suggestions.
          </Option>
        ) : hasNoResults ? (
          <Option key="__empty__" disabled value="">
            {noSuggestionMessage}
          </Option>
        ) : (
          filteredItems.map((item) => (
            <Option
              key={item.searchValue}
              value={item.outputValue}
              text={item.outputValue}
            >
              {getHighlightedText(item.displayValue, inputValue, styles.highlight)}
            </Option>
          ))
        )}
      </Combobox>

      {/*
        Indeterminate progress bar below the Combobox — visible even before the
        dropdown is opened, giving immediate feedback that a fetch is in flight.
      */}
      {isLoading && <ProgressBar className={styles.progressBar} />}

      {/*
        Error detail and Retry button rendered below the Combobox so they are
        always visible (not hidden behind a dropdown that might not be open).
        The dropdown itself shows "Failed to load options." as a disabled hint.
      */}
      {errorMessage && (
        <div className={styles.errorWrapper}>
          <span>{errorMessage}</span>
          {onRetry && (
            <button
              type="button"
              className={styles.retryButton}
              onClick={onRetry}
            >
              Retry
            </button>
          )}
        </div>
      )}
    </div>
  );
}
