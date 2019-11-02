import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class XrmMetadataAutoComplete implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	//private _labelElement : HTMLLabelElement;
	private _divContainer : HTMLDivElement;

	// input element that is used to create the autocomplete
	private _inputElement : HTMLInputElement;

	//Datalist element
	private _datalistElement : HTMLDataListElement;

	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;

	private _currentValue : string;

	 // PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;

	// Event Handler 'refreshData' reference
	private _refreshData: EventListenerOrEventListenerObject;

	private _autoCompleteValues: string[] = [];

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
		this._refreshData = this.refreshData.bind(this);
		
		var autoCompleteListUniqueId = this.uuidv4();

		// creating HTML elements for the input type list and binding it to the function which refreshes the control data
		this._inputElement = document.createElement("input");
		this._inputElement.setAttribute("list", "dataList" + autoCompleteListUniqueId);
        this._inputElement.setAttribute("name", "autoComplete");
		this._inputElement.addEventListener("input", this._refreshData);
		this._inputElement.setAttribute("value", this._context.parameters.selectedValue === undefined  || this._context.parameters.selectedValue.raw === null ? "" :  String(this._context.parameters.selectedValue.raw));
        
        // creating HTML elements for data list 
        this._datalistElement = document.createElement("datalist");
		this._datalistElement.setAttribute("id", "dataList" + autoCompleteListUniqueId);

		//Create options for data list
        var optionsHtml = "";
        var optionsHtmlArray = new Array();
       
		var metadataType = this._context.parameters.autoCompleteMetaDataType.raw;
		var relatedEntityName =  this._context.parameters.relatedEntity === undefined ? null : this._context.parameters.relatedEntity.raw;

		var filterEntityFieldByEntitiesAssociatedTo = this._context.parameters.filterEntityFieldByEntitiesAssociatedTo === undefined ? null : this._context.parameters.filterEntityFieldByEntitiesAssociatedTo.raw;

		var webApiUrl = "/api/data/v9.0/EntityDefinitions";

		var namefield = "LogicalName";
		var idField = "MetadataId";

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
			default:{
				this._divContainer.innerHTML = "Undefined metadatatype";
				break;
			}

		}

		//@ts-ignore
		const serverUrl = Xrm.Page.context.getClientUrl();
		let req = new XMLHttpRequest();
		req.open("GET", serverUrl + webApiUrl, true);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.onreadystatechange = () => {
			if (req.readyState == 4 /* complete */) {
				req.onreadystatechange = null; /* avoid potential memory leak issues.*/

				if (req.status == 200) {
					var data = JSON.parse(req.response);

					var results: Array<string>;
					results = new Array();
					var dataJson = data.value;

					if (metadataType === "Lookup"){
						dataJson = data.Attributes;
					}

					if (dataJson != null && dataJson.length > 0) {
						for (let i = 0; i < dataJson.length; i++) {
							
							if (metadataType  !== "SystemViews"){
							if (!this.ExistsinArray(results, dataJson[i][namefield])){
								results.push(dataJson[i][namefield]);
							}
							}
							else{
								results.push(dataJson[i][namefield] + "-CRMID-" + dataJson[i][idField]);	
							}
							
						}
					}

					this._autoCompleteValues = results.sort();
						
					for (var i = 0; i < this._autoCompleteValues.length; ++i) {
						optionsHtmlArray.push('<option value="');
						optionsHtmlArray.push(this._autoCompleteValues[i].toString());
						optionsHtmlArray.push('" />');
					}
					optionsHtml = optionsHtmlArray.join("");

				//@ts-ignore 
				this._datalistElement.innerHTML = optionsHtml;					

				} else {
					var error = JSON.parse(req.response).error;
					console.log(error.message);
				}
			}
		};
		req.send();
                        
        // appending the HTML elements to the control's HTML container element.
        //Add input element
        this._divContainer.appendChild(this._inputElement);

        //Add datalist element
        this._divContainer.appendChild(this._datalistElement);
		container.appendChild(this._divContainer);
	}

	/**
	 * Updates the values to the internal value variable we are storing and also updates the html label that displays the value
	 * @param evt : The "Input Properties" containing the parameters, control metadata and interface functions
	 */
	public refreshData(evt: Event): void {
		this._currentValue = (this._inputElement.value as any) as string;
		this._notifyOutputChanged();
	}



	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
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
		this._inputElement.removeEventListener("input", this._refreshData);
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
}