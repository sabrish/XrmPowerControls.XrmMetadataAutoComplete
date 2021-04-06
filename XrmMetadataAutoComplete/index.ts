import * as React from "react";
import * as ReactDOM from "react-dom";
import { ITag } from 'office-ui-fabric-react/lib/Pickers';
import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { ReactSearchBoxV2, IProps } from './Components/ReactSearchBox';
import { ISuggestionItem } from './Components/Autocomplete';


export class XrmMetadataAutoComplete implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	entity:string="";
	private firstRun:Boolean =true;
	/**
     * Selected items cache.
     */
	 public selectedItems: ITag[];

	
	//private _labelElement : HTMLLabelElement;
	private _divContainer : HTMLDivElement;

	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;

	private _currentValue : string;

	private _json:string|null;
	private _noSuggestions: string = "No data";
	private _searchTitle ="---";

	 // PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;

	// Event Handler 'refreshData' reference
	private _refreshData: EventListenerOrEventListenerObject;

	private _autoCompleteValues: ISuggestionItem[];

	private props:IProps = { value:"", json:[], onResult: this.notifyChange.bind(this),  noSuggestionMessage:this._noSuggestions, searchTitle: this._searchTitle };

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		// Add control initialization code

		this._context = context;
		this._divContainer = document.createElement("div");
		this._notifyOutputChanged = notifyOutputChanged;
		var metadataType = this._context.parameters.autoCompleteMetaDataType.raw;
		this.props.value = this._context.parameters.selectedValue.raw || "";
		var relatedEntityName =  this._context.parameters.relatedEntity === undefined ? null : this._context.parameters.relatedEntity.raw;

		var filterEntityFieldByEntitiesAssociatedTo = this._context.parameters.filterEntityFieldByEntitiesAssociatedTo === undefined ? null : this._context.parameters.filterEntityFieldByEntitiesAssociatedTo.raw;

		var webApiUrl = "/api/data/v9.0/EntityDefinitions";

		var namefield = "LogicalName";
		var idField = "MetadataId";
       
		if(metadataType == "Entity" && this.firstRun)
		{
			this.firstRun == false;
			this.PopulateDropDown(metadataType,filterEntityFieldByEntitiesAssociatedTo,webApiUrl,namefield,idField,relatedEntityName);
		}
		
		container.appendChild(this._divContainer);
	}

	notifyChange(value:string)
	{
		this._currentValue = value;
		this._notifyOutputChanged();

	}



	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		var metadataType = this._context.parameters.autoCompleteMetaDataType.raw;
		this.props.value = this._context.parameters.selectedValue.raw || "";
		var relatedEntityName =  this._context.parameters.relatedEntity === undefined ? null : this._context.parameters.relatedEntity.raw;

		var filterEntityFieldByEntitiesAssociatedTo = this._context.parameters.filterEntityFieldByEntitiesAssociatedTo === undefined ? null : this._context.parameters.filterEntityFieldByEntitiesAssociatedTo.raw;

		var webApiUrl = "/api/data/v9.0/EntityDefinitions";

		var namefield = "LogicalName";
		var idField = "MetadataId";
		if(metadataType != "Entity" && relatedEntityName != this.entity && relatedEntityName!=null)
		{
			this.entity = relatedEntityName;
			ReactDOM.unmountComponentAtNode(this._divContainer);
			this.PopulateDropDown(metadataType,filterEntityFieldByEntitiesAssociatedTo,webApiUrl,namefield,idField,relatedEntityName);
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		let result = {
            selectedValue: this._currentValue
		};
		
		return result;
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
		ReactDOM.unmountComponentAtNode(this._divContainer);
	}

	/** 
	 	* method to gnenerate unique ids in javassript so that I can have multiple PCF controls on the same form
		* https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
	*/
	private uuidv4():string {
		//@ts-ignore
		return ([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g, c =>
			(c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
		);
	}

	private ExistsinArray(data: Array<any>, stringToSearch: string) {
		//First sort the array and then run through it and then see if the next (or previous) index is the same as the current. 
		const sortedArr = data.slice().sort();
		for (var i = 0; i < data.length - 1; i++) {
			if (sortedArr[i] === stringToSearch) {
				return true;
			}
		}
		return false;
	}

	private async PopulateDropDown(metadataType:string,filterEntityFieldByEntitiesAssociatedTo:any,webApiUrl:string,namefield:string,idField:string,relatedEntityName:any)
	{
		switch(metadataType){
			case "Entity":{
				if (filterEntityFieldByEntitiesAssociatedTo == undefined || filterEntityFieldByEntitiesAssociatedTo == null)
				{
					webApiUrl = this.GetEntitiesUrl();
					namefield = "LogicalName";
					idField = "MetadataId";
				}
				else{
					webApiUrl = this.GetOneToManyRelationshipMetadataWithParamsUrl(String(filterEntityFieldByEntitiesAssociatedTo), "$select=ReferencingEntity,MetadataId");
					namefield = "ReferencingEntity";
					idField = "MetadataId";
				}

				
				break;
			}
			case "Attributes":{
				if (relatedEntityName != null){
					webApiUrl = this.GetAttributesforEntityUrl(relatedEntityName);
					namefield = "LogicalName";
					idField = "MetadataId";
				}
				else{
					this._divContainer.innerHTML = "Please provide the field which records the entity name or enter the entity name in the RelatedEntity property";
				}

				
				break;
			}
			case "Lookup":{
				if (relatedEntityName != null){
					webApiUrl = this.GetCustomerOrLookupAttributesforEntityUrl(relatedEntityName);
					namefield = "LogicalName";
					idField = "MetadataId";
				}
				else{
					this._divContainer.innerHTML = "Please provide the field which records the entity name or enter the entity name in the RelatedEntity property";
				}

				
				break;
			}
			case "SystemViews":{
				if (relatedEntityName != null){
					webApiUrl = this.GetSavedViewsForEntityUrl(relatedEntityName);
					namefield = "name";
					idField = "savedqueryid";
				}
				else{
					this._divContainer.innerHTML = "Please provide the field which records the entity name or enter the entity name in the RelatedEntity property";
				}
				
				break;
			}
			case "BusinessProcessFlows":{
				if (relatedEntityName != null){
					webApiUrl = this.GetBusinessProcessFlowsUrl(relatedEntityName);
					namefield="name";
					idField="workflowid";
				}
				
				break;
			}
			default:{
				this._divContainer.innerHTML = "Undefined metadatatype";
				break;
			}

		}

		//@ts-ignore
		const serverUrl = Xrm.Page.context.getClientUrl();
		var apiRequestUrl = serverUrl + webApiUrl;
		
		var results: ISuggestionItem[];
		results = [];
		var data = await this.getXrmMetaData(apiRequestUrl);
		var dataJson = data.value;
		if (metadataType === "Lookup"){
			dataJson = data.Attributes;
		}

		if (dataJson != null && dataJson.length > 0) {
			for (let i = 0; i < dataJson.length; i++) {
				
				if (!this.ExistsinArray(results, dataJson[i][namefield])){
					if (metadataType  == "SystemViews"){
						results.push({ key: i, displayValue: dataJson[i][namefield] + " (" + dataJson[i][idField]+")",searchValue:dataJson[i][namefield] + "-CRMID-" + dataJson[i][idField] });
					}
					else if(metadataType  == "BusinessProcessFlows")
					{
						results.push({ key: i, displayValue: dataJson[i][namefield],searchValue:dataJson[i][namefield] });
					}
					else
					{
						if(dataJson[i]["DisplayName"]["LocalizedLabels"].length>0)
						{
							results.push({ key: i, displayValue: dataJson[i]["DisplayName"]["LocalizedLabels"][0]["Label"]+" ("+dataJson[i][namefield]+")",searchValue:dataJson[i][namefield] });
						}
						else
						{
							results.push({ key: i, displayValue: dataJson[i][namefield],searchValue:dataJson[i][namefield] });
						}
					}
				}
				
			}
		}

		this._autoCompleteValues = results.sort((a, b) => (a.key > b.key) ? 1 : -1);
		this.props.json = this._autoCompleteValues;
		let obj = this.props.json.find((o, i) => {
			if (o.searchValue === this.props.value) {
				return true; // stop searching
			}
		});

		if(!obj)
		{
			this.props.value = "";
		}

		ReactDOM.render(
			React.createElement(ReactSearchBoxV2,this.props)
			, this._divContainer
		);			
	}

	private async getXrmMetaData(webApiUrl:string):Promise<any> {
		const response = await fetch(webApiUrl);
		const body = await response.json();
  		return body;
		
	}

	

	private GetEntitiesUrl() :string
	{
		return "/api/data/v9.0/EntityDefinitions";
	}

	private GetSavedViewsForEntityUrl(entitylogicalname: string) :string
	{
		return "/api/data/v9.0/savedqueries?$filter=returnedtypecode eq '" + entitylogicalname + "'";
	}

	private GetCustomerOrLookupAttributesforEntityUrl(entitylogicalname: string) :string
	{
		 return "/api/data/v9.0/EntityDefinitions(LogicalName='" +
			entitylogicalname +
			"')?$expand=Attributes($filter=AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Lookup' or AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Customer')";
	}

	private GetAttributesforEntityUrl(entitylogicalname: string): string
	{
		return "/api/data/v9.0/EntityDefinitions(LogicalName='" + entitylogicalname + "')/Attributes";
	}

	private  GetOneToManyRelationshipMetadataWithParamsUrl(entitylogicalname: string, paramstring: string ) :string
	{
		return "/api/data/v9.0/EntityDefinitions(LogicalName='" + entitylogicalname + "')/OneToManyRelationships?" + paramstring;
	}

	private GetBusinessProcessFlowsUrl(entitylogicalname: string): string
	{
		return "/api/data/v9.0/workflows?$filter=category eq 4 and primaryentity eq '"+entitylogicalname+"'";
	}
}