import { DefaultColors } from './styles/colors';
import { mergeStyleSets } from '@uifabric/styling';
export const CalloutStyle = () => {
  return {  width: '100%' };
};
export const AutocompleteStyles = () => {
  return ({
    marginTop: '10px', marginBottom: '20px', width:'100%', display: 'inline-block'
  });
};
export const SuggestionListStyle = () => {
  return ({ padding: '4px 16px', fontSize: '14px', cursor: 'default' });
};
export const SuggestionListItemStyle = mergeStyleSets({
  root: {
    selectors: {
     
      '&:focus': {
        //backgroundColor: DefaultColors.Item.ListItemHoverBackgroundColor
      backgroundColor: "#f3f2f1",
      color: "black",
      outline: "none"
      }
    }
  }
});