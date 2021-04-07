import * as React from 'react';
import { Pivot, PivotItem, Fabric, initializeIcons,IColumn,Selection } from 'office-ui-fabric-react';
import { Autocomplete, ISuggestionItem } from './Autocomplete';


export interface IDetailsListCompactItem {
  key: number;
  displayValue: string;
  searchValue: string;
}

initializeIcons();

export interface IState {
	value:string;
  json:Array<ISuggestionItem>;
}

export interface IProps {
	value:string;
  json:Array<ISuggestionItem>;
  onResult: (value:string) =>void;
  onChange: (value?:string) => void;
  noSuggestionMessage:string;
  searchTitle:string;
}

              


export class ReactSearchBoxV2 extends React.Component<IProps, IState> {
  private _selection: Selection;
   
  private favcolumns: IColumn[];

  constructor(props: Readonly<IProps>) {
      super(props);
      
      this.state = {value:props.value, json:props.json};
      this.entitySelectHandler = this.entitySelectHandler.bind(this);
      this.searchTextandler = this.searchTextandler.bind(this);
      this.onChangeHandler = this.onChangeHandler.bind(this);
      //this._noSuggestionMessage = props.noSuggestionMessage;
      

       //Funciones para favoritos
      this._selection = new Selection({
        onSelectionChanged: () => {
          this.setState({ value: this.getSelectionDetails() })
          this.props.onResult(this.getSelectionDetails());
        },
      });

    }
  
    //Funciones para el seartchbox + suggestbox
  entitySelectHandler = (item: ISuggestionItem): void => {
    this.setState({value: item.searchValue});
    this.props.onResult(item.searchValue);
    
  }
  
  searchTextandler = (item: string): void => {
   this.setState({value: item as string});
  }

  onChangeHandler = (item?: string): void => {
      console.log("onchange triggered");
    this.setState({value: item as string});
    this.props.onChange(item);
  }
  
  //Funciones para el favorito
  getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    
    if (selectionCount ==1) return (this._selection.getSelection()[0] as IDetailsListCompactItem).searchValue;
    return "";
  }
 
	 
  render() {

    return (
       <Fabric>

          <Autocomplete
            items={this.state.json}
            searchTitle={this.props.searchTitle}
            suggestionCallback={this.entitySelectHandler}
            searchCallback={this.searchTextandler}
            onChangeCallback = {this.onChangeHandler}
            value = {this.state.value}
            noSuggestionsMessage = {this.props.noSuggestionMessage}
          />
          
       </Fabric>
    );
  }
}