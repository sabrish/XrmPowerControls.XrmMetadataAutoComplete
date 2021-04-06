import * as React from 'react';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { ComboBox, IComboBox, IComboBoxOption, IComboBoxOptionStyles, IComboBoxStyles } from 'office-ui-fabric-react';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

export interface IComboOptions {
    items: any[],
    id:string,
    selectedValue?:string
    dropdownChangedValue:(newValue:string) => void;
  }

const stackTokens: IStackTokens = { childrenGap: 20 };
const comboBoxStyles: Partial<IComboBoxOptionStyles> = { root: { alignItems: 'bottom' } };

export class ComboBoxExample extends React.Component<IComboOptions> {
    constructor(props: Readonly<IComboOptions>) {
        super(props);
        this.state = { Counter: 0 };
        
    }
    
    render() {
        return (
          
            <Stack tokens={stackTokens}>
            <ComboBox  
                comboBoxOptionStyles={comboBoxStyles}
                defaultSelectedKey = {this.props.selectedValue}
                label='' 
                id={this.props.id}  
                ariaLabel='Basic ComboBox example' 
                allowFreeform={ true } 
                autoComplete='on'  
                options = {this.props.items }
                onChange={this._onChange} 
                />
            </Stack>
        );
    }

    private _onChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
        if(option){
          const selectedKey: string = option.key as string
          if (this.props.dropdownChangedValue) {
            this.props.dropdownChangedValue(selectedKey);
          }
        }
        else{
           this.props.dropdownChangedValue("");
        }  
      };

     
};
