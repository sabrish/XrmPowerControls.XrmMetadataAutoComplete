# XrmPowerControls.XrmMetadataAutoComplete
PCF control for metadata autocomplete in D365, CDS Environments

This control can be used to provide autocomplete functionality for CRM metadata and store them in text fields.

### Usage

![alt text](https://github.com/sabrish/XrmPowerControls.XrmMetadataAutoComplete/blob/master/XrmPowerControls.XrmMetadataAutoComplete.gif?raw=true)

Parameters used in the control are:
* **AutoCompleteMetaDataType**: This is of type Enum. The suppored values are
  * _Entity_
  * _Attributes_
  * _Lookup_
  * _SystemViews_ - Will have [ViewName]-{CRMID}-[View GUID] format to cope with views with same name
* **Filter Entity Field By Entities Associated To**: Only applies if the AutoCompleteMetaDataType is set to Entity. It will filter entity field auto complete list By entities associated to entity entered in this property.
* **Related Entity**: Only applies if the AutoCompleteMetaDataType is set to anything other than Entity. It will filter the autocomplete list to contain only metatadata values related to this entity.



__Note and Thanks to the PCF Autocompelte project by Sriram Balaji - https://github.com/srirambalajigit as I used his project 
https://github.com/srirambalajigit/PCFControls/tree/master/Autocomplete/Autocomplete as a reference when building my first PCF control.__
