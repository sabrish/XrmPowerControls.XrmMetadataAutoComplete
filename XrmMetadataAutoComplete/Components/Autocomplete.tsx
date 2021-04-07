import * as React from 'react';
import { SearchBox, Callout, List, DirectionalHint, Stack, IStackTokens } from 'office-ui-fabric-react/lib/';
import {
  CalloutStyle, AutocompleteStyles, SuggestionListStyle, 
  SuggestionListItemStyle
} from './AutoComplete.style';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';

const stackTokens: Partial<IStackTokens> = { childrenGap: 20 };

export interface IAutocompleteProps {
  items: ISuggestionItem[];
  searchTitle?: string;
  suggestionCallback: (item: ISuggestionItem) => void;
  searchCallback: (item: string) => void;
  value:string;
  noSuggestionsMessage:string;
}
export interface IAutocompleteState {
  isSuggestionDisabled: boolean;
  searchText: string;
  value:string;
}
export interface ISuggestionItem {
  key: number;
  displayValue: string;
  searchValue: string;
  type?: string;
  tag?: any; 
}

const KeyCodes = {
  tab: 9 as 9,
  enter: 13 as 13,
  left: 37 as 37,
  up: 38 as 38,
  right: 39 as 39,
  down: 40 as 40,
}

 

type ISearchSuggestionsProps = IAutocompleteProps;

export class Autocomplete  extends React.Component<ISearchSuggestionsProps, IAutocompleteState> {

private _searchContainerRef = React.createRef<HTMLDivElement>();

  constructor(props: ISearchSuggestionsProps) {
    super(props);
    this.state = {
      isSuggestionDisabled: false,
      searchText: '',
      value: props.value
    };
  }
  protected getComponentName(): string {
    return 'SearchSuggestions';
  }
  handleClick = (item: ISuggestionItem) => {
    this.props.suggestionCallback(item);
    this.setState({ value:item.searchValue})
    
  }
  render() {
    return (
      this.renderSearch()
    );
  }
  private renderSearch = () => {
    let showHide:boolean = this.state.value=="" ? true : false;
     
    if(showHide) {
    return (
     
      <div ref={this._searchContainerRef} style={AutocompleteStyles()} onKeyDown={this.onKeyDown}>
        <Stack tokens={stackTokens}>
          <SearchBox
            id={'SuggestionSearchBox'}
            placeholder={this.props.searchTitle}
            onSearch={newValue => this.onSearch(newValue)}
            onClick={newSearchText => { this.showSuggestionCallOut();}}
            onChange={newSearchText => {
              newSearchText && newSearchText.currentTarget.value.trim() !== '' ? this.showSuggestionCallOut() : this.hideSuggestionCallOut();
              this.setState({ searchText: (newSearchText && newSearchText.currentTarget.value) as string });     
            }}
          onClear={(ev:any)=>this.setState({ value:"", searchText:"", isSuggestionDisabled:true})}
          disableAnimation
          />
        </Stack>
        {this.renderSuggestions()}         
      </div>
    );
        }
        else
        return (
          <div ref={this._searchContainerRef} style={AutocompleteStyles()} onKeyDown={this.onKeyDown}>
            <div ref={this._searchContainerRef}>
            <Stack tokens={stackTokens}>
              <SearchBox
                id={'SuggestionSearchBox'}
                placeholder={this.props.searchTitle}
                onClick={newSearchText => { this.showSuggestionCallOut();}}
                onSearch={newValue => this.onSearch(newValue)}
                onChange={newSearchText => {
                  newSearchText && newSearchText.currentTarget.value.trim() !== '' ? this.showSuggestionCallOut() : this.hideSuggestionCallOut();
                  this.setState({ searchText: (newSearchText && newSearchText.currentTarget.value) as string, value:"" });     
                }}
              onClear={(ev:any)=>this.setState({ value:"", searchText:"", isSuggestionDisabled:true})}      
              value = {this.state.value}
              disableAnimation
              /> 
              </Stack>      
            </div>
        </div>

        )
  }

  private onSearch(enteredEntityValue: string) {
    this.props.searchCallback(enteredEntityValue.trim());
  }
  private renderSuggestions = () => {
    return (
      <Callout id='SuggestionContainer'
        ariaLabelledBy={'callout-suggestions'}
        gapSpace={2}
        coverTarget={false}
        alignTargetEdge={true}
        onDismiss={ev => this.hideSuggestionCallOut()}
        setInitialFocus={false}
        hidden={!this.state.isSuggestionDisabled}
        calloutMaxHeight={300}
        style={CalloutStyle()}
        target={this._searchContainerRef.current}
        directionalHint={DirectionalHint.bottomLeftEdge}
        isBeakVisible={false}
        directionalHintFixed={true}
      >
        {this.renderSuggestionList()}
      </Callout >
    );
  }
  private renderSuggestionList = () => {
    if(this.state.searchText != undefined)
    {
      return (
        
          <List id='SearchList' tabIndex={0}
            items={this.suggestedTagsFiltered(this.props.items)}
            onRenderCell={this.onRenderCell}
          />
        
      );
    }
  }
  private onRenderCell = (item: any) => {
    if (item.key !== -1) {
      return (
        <div key={item.key}
          className={SuggestionListItemStyle.root}
          data-is-focusable={true}
          onKeyDown={(ev: React.KeyboardEvent<HTMLElement>) => this.handleListItemKeyDown(ev, item)}>
          <div id={'link' + item.key}
            style={SuggestionListStyle()}
            onClick={() => this.handleClick(item)}>
            {item.displayValue}
          </div>
        </div>
      );
    } else {
      return (
        <div key={item.key} data-is-focusable={true}>
          {item.displayValue}
        </div>
      );
    }
  }

  private showSuggestionCallOut() {
    this.setState({ isSuggestionDisabled: true });
  }
  private hideSuggestionCallOut() {
    this.setState({ isSuggestionDisabled: false });
  }
  private suggestedTagsFiltered = (list: ISuggestionItem[]) => {
    let suggestedTags = list.filter(tag => this.state.searchText.toLowerCase() == "" ||
      tag.searchValue.toLowerCase().includes(this.state.searchText.toLowerCase()));
    suggestedTags = suggestedTags.sort((a, b) => a.searchValue.localeCompare(b.searchValue));
    if (suggestedTags.length === 0) {
      suggestedTags = [{ key: -1, displayValue: this.props.noSuggestionsMessage, searchValue: '' }];
    }
    return suggestedTags;
  }
  protected handleListItemKeyDown = (ev: React.KeyboardEvent<HTMLElement>, item: ISuggestionItem): void => {
    const keyCode = ev.which;
    switch (keyCode) {
      case KeyCodes.enter:
        this.handleClick(item);
        break;
    }
  };
  protected onKeyDown = (ev: React.KeyboardEvent<HTMLElement>): void => {
    const keyCode = ev.which;
    switch (keyCode) {
      case KeyCodes.down:
        let el: any = window.document.querySelector("#SearchList");
        el.focus();
        break;
    }
  };
}