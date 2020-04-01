import { DynamicProperty } from '@microsoft/sp-component-base';
import { ISortFieldConfiguration, ISortFieldDirection } from '../../models/ISortFieldConfiguration';
import ISortableFieldConfiguration from '../../models/ISortableFieldConfiguration';
import { ISynonymFieldConfiguration } from '../../models/ISynonymFieldConfiguration';
import { BaseDataSource } from '../BaseDataSource';
import { IPropertyPaneGroup, IPropertyPaneDropdownOption, IPropertyPaneConditionalGroup, IPropertyPaneField, PropertyPaneHorizontalRule, PropertyPaneCheckbox, PropertyPaneTextField, PropertyPaneDynamicFieldSet, PropertyPaneDynamicField, DynamicDataSharedDepth, PropertyPaneToggle, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import ISearchService from '../../services/SearchService/ISearchService';
import { ServiceScope } from '@microsoft/sp-core-library';
import SearchService from '../../services/SearchService/SearchService';
import { SearchComponentType } from '../../models/SearchComponentType';
import * as React from 'react';
import { SearchManagedProperties, ISearchManagedPropertiesProps } from '../../controls/SearchManagedProperties/SearchManagedProperties';
import { PropertyPaneSearchManagedProperties } from '../../controls/PropertyPaneSearchManagedProperties/PropertyPaneSearchManagedProperties';
import { IDropdownOption, IComboBoxOption } from 'office-ui-fabric-react';
import { sortBy, cloneDeep } from '@microsoft/sp-lodash-subset';
import { IDataContext } from '../IDataContext';
import { IDataSourceData } from '../IDataSourceData';
import * as dataSourceStrings from 'SharePointDataSourceStrings';
import IDynamicDataService from '../../services/DynamicDataService/IDynamicDataService';
import { DynamicDataService } from '../../services/DynamicDataService/DynamicDataService';
import IRefinerSourceData from '../../models/IRefinerSourceData';
import ISearchVerticalSourceData from '../../models/ISearchVerticalSourceData';
import { ISearchVerticalInformation } from '../../models/ISearchResult';

/**
 * SharePoint search data source property pane properties
 */
export interface ISharePointSearchDataSourceProperties {

    /**
     * The query text
     */
    queryKeywords: DynamicProperty<string>;

    /**
     * The configured default query text to use with conencted to a dynamic source
     */
    defaultSearchQuery: string;

    /**
     * Flag indicating if a default search query should be use on first load
     */
    useDefaultSearchQuery: boolean;

    /**
     * The search query tempalte
     */
    queryTemplate: string;

    /**
     * The SharePoint result source ID
     */
    resultSourceId: string;

    /**
     * The initial sort fields and directions for the search request
     */
    sortList: ISortFieldConfiguration[];

    /**
     * Flag indicating if SharePoint query rules should be used
     */
    enableQueryRules: boolean;

    /**
     * Flag indication if the OneDrive for Buisiness should be included in the results
     */
    includeOneDriveResults: boolean;

    /**
     * The search selected managed properties
     */
    selectedProperties: string;

    /**
     * The sortable fields to display in the UI
     */
    sortableFields: ISortableFieldConfiguration[];

    /**
     * Flag indicating if the taxonomy based values should be translated according to the current UI language
     */
    enableLocalization: boolean;

    /**
     * The synonym list
     */
    synonymList: ISynonymFieldConfiguration[];

    /**
     * The search query language to use
     */
    searchQueryLanguage: number;

    /**
     * The search refinement filters
     */
    refinementFilters: string;

    /**
     * Flag indicating if the source connects to refiners
     */
    useRefiners: boolean;

    /**
     * Flag indicating if the source connects to search verticals
     */
    useSearchVerticals: boolean;

    /**
     * Dynamic data references
     */
    refinerDataSourceReference: string;
    searchVerticalDataSourceReference: string;
}

export class SharePointSearchDataSource extends BaseDataSource<ISharePointSearchDataSourceProperties> {

    private _propertyFieldCollectionData = null;
    private _customCollectionFieldType = null;

    private _availableLanguages: IPropertyPaneDropdownOption[];

    private _searchService: ISearchService;
    private _dynamicDataService: IDynamicDataService;

    /**
     * The list of available managed managed properties (managed globally for all property pane fiels if needed)
     */
    private _availableManagedProperties: IComboBoxOption[];

    private _verticalsInformation: ISearchVerticalInformation[];

    private _refinerSourceData: DynamicProperty<IRefinerSourceData>;
    private _searchVerticalSourceData: DynamicProperty<ISearchVerticalSourceData>;

    public constructor(serviceScope: ServiceScope) {
        super(serviceScope);

        this._availableManagedProperties = [];

        this._onUpdateAvailableProperties = this._onUpdateAvailableProperties.bind(this);

        serviceScope.whenFinished(() => {
            this._searchService = serviceScope.consume<ISearchService>(SearchService.ServiceKey);
            this._dynamicDataService = serviceScope.consume<IDynamicDataService>(DynamicDataService.ServiceKey);
        }); 
    }

    public onInit() {
        this.ensureDataSourceConnection();
    }

    public getData(dataContext?: IDataContext): Promise<IDataSourceData> {
        throw new Error("Method not implemented.");
    }
    
    public getItemCount(): number {
        throw new Error("Method not implemented.");
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] | IPropertyPaneConditionalGroup[] {

        return [
            {
                groupFields: this._getSearchSettingsFields(),
                isCollapsed: false,
                groupName: dataSourceStrings.PropertyPane.SearchSettingsGroupName
            }
        ]; 
    }

    public async onPropertyPaneConfigurationStart(): Promise<void> {
        await this.loadPropertyPaneResources();
    }
    
    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any) {
        // Bind connected data sources
        if (this.properties.refinerDataSourceReference || this.properties.searchVerticalDataSourceReference) {
            this.ensureDataSourceConnection();
        }

        if (propertyPath.localeCompare('useRefiners') === 0) {
            if (!this.properties.useRefiners) {
                this.properties.refinerDataSourceReference = undefined;
                this._refinerSourceData = undefined;
                this.context.dynamicDataSourceManager.notifyPropertyChanged(SearchComponentType.SearchResultsWebPart);
            }
        }

        if (propertyPath.localeCompare('useSearchVerticals') === 0) {

            if (!this.properties.useSearchVerticals) {
                this.properties.searchVerticalDataSourceReference = undefined;
                this._searchVerticalSourceData = undefined;
                this._verticalsInformation = [];
                this.context.dynamicDataSourceManager.notifyPropertyChanged(SearchComponentType.SearchResultsWebPart);
            }
        }

        if (propertyPath.localeCompare('searchVerticalDataSourceReference') === 0 || propertyPath.localeCompare('refinerDataSourceReference')) {
            this.context.dynamicDataSourceManager.notifyPropertyChanged(SearchComponentType.SearchResultsWebPart);
        }

        if (this.properties.enableLocalization) {

            let udpatedProperties: string[] = this.properties.selectedProperties.split(',');
            if (udpatedProperties.indexOf('UniqueID') === -1) {
                udpatedProperties.push('UniqueID');
            }

            // Add automatically the UniqueID managed property for subsequent queries
            this.properties.selectedProperties = udpatedProperties.join(',');
        }

        // clean out duplicate ones
        let allProps = this.properties.selectedProperties.split(',');
        allProps = allProps.filter((item, index) => {
            return allProps.indexOf(item) === index;
        });
        this.properties.selectedProperties = allProps.join(',');
    }

    private async loadPropertyPaneResources(): Promise<void> {

        const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
            /* webpackChunkName: 'search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
        );
        this._propertyFieldCollectionData = PropertyFieldCollectionData;
        this._customCollectionFieldType = CustomCollectionFieldType;

        if (this._availableLanguages.length == 0) {
            const languages = await this._searchService.getAvailableQueryLanguages();

            this._availableLanguages = languages.map(language => {
                return {
                    key: language.Lcid,
                    text: `${language.DisplayName} (${language.Lcid})`
                };
            });
        }
    }

    /**
     * Determines the group fields for the search settings options inside the property pane
     */
    private _getSearchSettingsFields(): IPropertyPaneField<any>[] {

        // Get available data source Web Parts on the page
        const refinerWebParts = this._dynamicDataService.getAvailableDataSourcesByType(SearchComponentType.RefinersWebPart);
        const searchVerticalsWebParts = this._dynamicDataService.getAvailableDataSourcesByType(SearchComponentType.SearchVerticalsWebPart);

        let useRefiners = this.properties.useRefiners;
        let useSearchVerticals = this.properties.useSearchVerticals;

        if (this.properties.useRefiners && refinerWebParts.length === 0) {
            useRefiners = false;
        }

        if (this.properties.useSearchVerticals && searchVerticalsWebParts.length === 0) {
            useSearchVerticals = false;
        }

        // Sets up search settings fields
        const searchSettingsFields: IPropertyPaneField<any>[] = [
            PropertyPaneTextField('queryTemplate', {
                label: dataSourceStrings.PropertyPane.QueryTemplateFieldLabel,
                value: this.properties.queryTemplate,
                disabled: this.properties.searchVerticalDataSourceReference ? true : false,
                multiline: true,
                resizable: true,
                placeholder: dataSourceStrings.PropertyPane.SearchQueryPlaceHolderText,
                deferredValidationTime: 300
            }),
            PropertyPaneTextField('resultSourceId', {
                label: dataSourceStrings.PropertyPane.ResultSourceIdLabel,
                multiline: false,
                onGetErrorMessage: this._validateSourceId.bind(this),
                deferredValidationTime: 300
            }),
            this._propertyFieldCollectionData('sortList', {
                manageBtnLabel: dataSourceStrings.PropertyPane.Sort.EditSortLabel,
                key: 'sortList',
                enableSorting: true,
                panelHeader: dataSourceStrings.PropertyPane.Sort.EditSortLabel,
                panelDescription: dataSourceStrings.PropertyPane.Sort.SortListDescription,
                label: dataSourceStrings.PropertyPane.Sort.SortPropertyPaneFieldLabel,
                value: this.properties.sortList,
                fields: [
                    {
                        id: 'sortField',
                        title: "Field name",
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {

                            // Need to specify a React key to avoid item duplication when adding a new row
                            return React.createElement("div", { key: `${field.id}-${itemId}` },
                                React.createElement(SearchManagedProperties, {
                                    defaultSelectedKey: item[field.id] ? item[field.id] : '',
                                    onUpdate: (newValue: any, isSortable: boolean) => {

                                        if (!isSortable) {
                                            onCustomFieldValidation(field.id, dataSourceStrings.PropertyPane.Sort.SortInvalidSortableFieldMessage);
                                        } else {
                                            onUpdate(field.id, newValue);
                                            onCustomFieldValidation(field.id, '');
                                        }
                                    },
                                    searchService: this._searchService,
                                    validateSortable: true,
                                    availableProperties: this._availableManagedProperties,
                                    onUpdateAvailableProperties: this._onUpdateAvailableProperties
                                } as ISearchManagedPropertiesProps));
                        }
                    },
                    {
                        id: 'sortDirection',
                        title: "Direction",
                        type: this._customCollectionFieldType.dropdown,
                        required: true,
                        options: [
                            {
                                key: ISortFieldDirection.Ascending,
                                text: dataSourceStrings.PropertyPane.Sort.SortDirectionAscendingLabel
                            },
                            {
                                key: ISortFieldDirection.Descending,
                                text: dataSourceStrings.PropertyPane.Sort.SortDirectionDescendingLabel
                            }
                        ]
                    }
                ]
            }),
            this._propertyFieldCollectionData('sortableFields', {
                manageBtnLabel: dataSourceStrings.PropertyPane.Sort.EditSortableFieldsLabel,
                key: 'sortableFields',
                enableSorting: true,
                panelHeader: dataSourceStrings.PropertyPane.Sort.EditSortableFieldsLabel,
                panelDescription: dataSourceStrings.PropertyPane.Sort.SortableFieldsDescription,
                label: dataSourceStrings.PropertyPane.Sort.SortableFieldsPropertyPaneField,
                value: this.properties.sortableFields,
                fields: [
                    {
                        id: 'sortField',
                        title: dataSourceStrings.PropertyPane.Sort.SortableFieldManagedPropertyField,
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            // Need to specify a React key to avoid item duplication when adding a new row
                            return React.createElement("div", { key: `${field.id}-${itemId}` },
                                React.createElement(SearchManagedProperties, {
                                    defaultSelectedKey: item[field.id] ? item[field.id] : '',
                                    onUpdate: (newValue: any, isSortable: boolean) => {

                                        if (!isSortable) {
                                            onCustomFieldValidation(field.id, dataSourceStrings.PropertyPane.Sort.SortInvalidSortableFieldMessage);
                                        } else {
                                            onUpdate(field.id, newValue);
                                            onCustomFieldValidation(field.id, '');
                                        }
                                    },
                                    searchService: this._searchService,
                                    validateSortable: true,
                                    availableProperties: this._availableManagedProperties,
                                    onUpdateAvailableProperties: this._onUpdateAvailableProperties
                                } as ISearchManagedPropertiesProps));
                        }
                    },
                    {
                        id: 'displayValue',
                        title: dataSourceStrings.PropertyPane.Sort.SortableFieldDisplayValueField,
                        type: this._customCollectionFieldType.string
                    },
                    {
                      id: 'sortDirection',
                      title: "Direction",
                      type: this._customCollectionFieldType.dropdown,
                      required: true,
                      options: [
                          {
                              key: ISortFieldDirection.Ascending,
                              text: dataSourceStrings.PropertyPane.Sort.SortDirectionAscendingLabel
                          },
                          {
                              key: ISortFieldDirection.Descending,
                              text: dataSourceStrings.PropertyPane.Sort.SortDirectionDescendingLabel
                          }
                      ]
                  }
                ]
            }),
            PropertyPaneToggle('useRefiners', {
                label: dataSourceStrings.PropertyPane.UseRefinersWebPartLabel,
                checked: useRefiners
            }),
            PropertyPaneToggle('useSearchVerticals', {
                label: dataSourceStrings.PropertyPane.ConnectToSearchVerticalsLabel,
                checked: useSearchVerticals
            }),
            PropertyPaneToggle('enableQueryRules', {
                label: dataSourceStrings.PropertyPane.EnableQueryRulesLabel,
                checked: this.properties.enableQueryRules,
            }),
            PropertyPaneToggle('includeOneDriveResults', {
                label: dataSourceStrings.PropertyPane.IncludeOneDriveResultsLabel,
                checked: this.properties.includeOneDriveResults,
            }),
            new PropertyPaneSearchManagedProperties('selectedProperties', {
                label: dataSourceStrings.PropertyPane.SelectedPropertiesFieldLabel,
                description: dataSourceStrings.PropertyPane.SelectedPropertiesFieldDescription,
                allowMultiSelect: true,
                availableProperties: this._availableManagedProperties,
                defaultSelectedKeys: this.properties.selectedProperties.split(","),
                onPropertyChange: (propertyPath: string, newValue: any) => {
                    this.properties[propertyPath] = newValue.join(',');
                    this.onPropertyUpdate(propertyPath, this.properties.selectedProperties, newValue);

                    // Refresh the WP with new selected properties
                    this.render();
                },
                onUpdateAvailableProperties: this._onUpdateAvailableProperties,
                searchService: this._searchService,
            }),
            PropertyPaneTextField('refinementFilters', {
                label: dataSourceStrings.PropertyPane.RefinementFilters,
                 multiline: true,
                 deferredValidationTime: 300
            }),
            PropertyPaneToggle('enableLocalization', {
                checked: this.properties.enableLocalization,
                label: dataSourceStrings.PropertyPane.EnableLocalizationLabel,
                onText: dataSourceStrings.PropertyPane.EnableLocalizationOnLabel,
                offText: dataSourceStrings.PropertyPane.EnableLocalizationOffLabel
            }),
            PropertyPaneDropdown('searchQueryLanguage', {
                label: dataSourceStrings.PropertyPane.QueryCultureLabel,
                options: [{
                    key: -1,
                    text: dataSourceStrings.PropertyPane.QueryCultureUseUiLanguageLabel
                } as IDropdownOption].concat(sortBy(this._availableLanguages, ['text'])),
                selectedKey: this.properties.searchQueryLanguage ? this.properties.searchQueryLanguage : 0
            }),
            this._propertyFieldCollectionData('synonymList', {
                manageBtnLabel: dataSourceStrings.PropertyPane.Synonyms.EditSynonymLabel,
                key: 'synonymList',
                enableSorting: false,
                panelHeader: dataSourceStrings.PropertyPane.Synonyms.EditSynonymLabel,
                panelDescription: dataSourceStrings.PropertyPane.Synonyms.SynonymListDescription,
                label: dataSourceStrings.PropertyPane.Synonyms.SynonymPropertyPanelFieldLabel,
                value: this.properties.synonymList,
                fields: [
                    {
                        id: 'Term',
                        title: dataSourceStrings.PropertyPane.Synonyms.SynonymListTerm,
                        type: this._customCollectionFieldType.string,
                        required: true,
                        placeholder: dataSourceStrings.PropertyPane.Synonyms.SynonymListTermExemple
                    },
                    {
                        id: 'Synonyms',
                        title: dataSourceStrings.PropertyPane.Synonyms.SynonymListSynonyms,
                        type: this._customCollectionFieldType.string,
                        required: true,
                        placeholder: dataSourceStrings.PropertyPane.Synonyms.SynonymListSynonymsExemple
                    },
                    {
                        id: 'TwoWays',
                        title: dataSourceStrings.PropertyPane.Synonyms.SynonymIsTwoWays,
                        type: this._customCollectionFieldType.boolean,
                        required: false
                    }
                ]
            })
        ];

        // Conditional fields for data sources
        if (this.properties.useRefiners) {

            searchSettingsFields.splice(5, 0,
                PropertyPaneDropdown('refinerDataSourceReference', {
                    options: this._dynamicDataService.getAvailableDataSourcesByType(SearchComponentType.RefinersWebPart),
                    label: dataSourceStrings.PropertyPane.UseRefinersFromComponentLabel
                }));
        }

        if (this.properties.useSearchVerticals) {
            searchSettingsFields.splice(this.properties.useRefiners ? 7 : 6, 0,
                PropertyPaneDropdown('searchVerticalDataSourceReference', {
                    options: this._dynamicDataService.getAvailableDataSourcesByType(SearchComponentType.SearchVerticalsWebPart),
                    label: "Use verticals from this component"
                }));
        }

        return searchSettingsFields;
    }

    /**
     * Ensures the result source id value is a valid GUID
     * @param value the result source id
     */
    private _validateSourceId(value: string): string {
        if (value.length > 0) {
            if (!(/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/).test(value)) {
                return dataSourceStrings.PropertyPane.InvalidResultSourceIdMessage;
            }
        }

        return '';
    }

    /**
     * Handler when the list of available managed properties is fetched by a property pane control¸or a field in a collection data control
     * @param properties the fetched properties
     */
    private _onUpdateAvailableProperties(properties: IComboBoxOption[]) {

        // Save the value in the root Web Part class to avoid fetching it again if the property list is requested again by any other property pane control
        this._availableManagedProperties = cloneDeep(properties);

        // Refresh all fields so other property controls can use the new list
        this.context.propertyPane.refresh();
        this.render();
    }

    /**
     * Make sure the dynamic property is correctly connected to the source if a search refiner component has been selected in options
     */
    private ensureDataSourceConnection() {

        // Refiner Web Part data source
        if (this.properties.refinerDataSourceReference) {

            if (!this._refinerSourceData) {
                this._refinerSourceData = new DynamicProperty<IRefinerSourceData>(this.context.dynamicDataProvider);
            }

            // Register the data source manually since we don't want user select properties manually
            this._refinerSourceData.setReference(this.properties.refinerDataSourceReference);
            this._refinerSourceData.register(this.render);

        } else {

            if (this._refinerSourceData) {
                this._refinerSourceData.unregister(this.render);
            }
        }

        // Search verticals Web Part data source
        if (this.properties.searchVerticalDataSourceReference) {

            if (!this._searchVerticalSourceData) {
                this._searchVerticalSourceData = new DynamicProperty<ISearchVerticalSourceData>(this.context.dynamicDataProvider);
            }

            // Register the data source manually since we don't want user select properties manually
            this._searchVerticalSourceData.setReference(this.properties.searchVerticalDataSourceReference);
            this._searchVerticalSourceData.register(this.render);

        } else {

            if (this._searchVerticalSourceData) {
                this._searchVerticalSourceData.unregister(this.render);
            }
        }
    }
}