import { makeStyles, tokens } from '@fluentui/react-components';

export const useAutocompleteStyles = makeStyles({
  /**
   * Outer wrapper around the Combobox.
   * width:100% ensures the control fills its PCF container.
   * No marginBottom — PCF form layout owns vertical spacing between fields.
   */
  container: {
    marginTop: tokens.spacingVerticalS,
    width: '100%',
  },

  /**
   * Applied to the Combobox root element so it stretches to fill the container.
   * Without this the Combobox renders at its intrinsic (narrow) width.
   */
  combobox: {
    width: '100%',
  },

  /**
   * Error row rendered below the Combobox when a fetch fails.
   * Always visible (not inside the dropdown) so the user sees it without
   * having to open the dropdown.
   */
  errorWrapper: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    paddingTop: tokens.spacingVerticalXXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteRedForeground1,
  },

  /** Inline retry link in the error state. */
  retryButton: {
    background: 'none',
    border: 'none',
    padding: 0,
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    fontSize: tokens.fontSizeBase200,
    fontFamily: tokens.fontFamilyBase,
    textDecoration: 'underline',
    flexShrink: 0,
    ':hover': {
      color: tokens.colorBrandForeground2Hover,
    },
  },

  /**
   * Indeterminate progress bar that appears below the Combobox while loading.
   * Visible even when the dropdown is closed, giving immediate fetch feedback.
   */
  progressBar: {
    marginTop: tokens.spacingVerticalXXS,
  },

  /**
   * Highlighted span wrapping query-matching portions of a suggestion's
   * displayValue. Uses a subtle brand tint so it works in both light and dark
   * modes without hardcoded colours.
   */
  highlight: {
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    borderRadius: tokens.borderRadiusSmall,
  },
});
