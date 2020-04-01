declare interface ISharePointDataSourceStrings {
    General: {
        EmptyFieldErrorMessage: string;
    },
    PropertyPane: {
        SearchSettingsGroupName: string;
        UseDefaultSearchQueryKeywordsFieldLabel: string;
        DefaultSearchQueryKeywordsFieldLabel: string;
        DefaultSearchQueryKeywordsFieldDescription: string;
        SearchQueryPlaceHolderText: string;
        UseDefaultSearchQueryKeywordsFieldLabel: string;
        DefaultSearchQueryKeywordsFieldLabel: string;
        SearchQueryKeywordsFieldLabel: string;
        SearchQueryKeywordsFieldDescription: string;
        QueryTemplateFieldLabel: string;
        ResultSourceIdLabel: string;
        Sort: {
            SortPropertyPaneFieldLabel
            SortListDescription: string;
            SortDirectionAscendingLabel:string;
            SortDirectionDescendingLabel:string;
            SortErrorMessage:string;
            SortPanelSortFieldLabel:string;
            SortPanelSortFieldAria:string;
            SortPanelSortFieldPlaceHolder:string;
            SortPanelSortDirectionLabel:string;
            SortableFieldsPropertyPaneField: string;
            SortableFieldsDescription: string;
            SortableFieldManagedPropertyField: string;
            SortableFieldDisplayValueField: string;
            EditSortableFieldsLabel: string;
            EditSortLabel: string;
            SortInvalidSortableFieldMessage: string;
        },
        UseRefinersWebPartLabel: string;
        UseRefinersFromComponentLabel: string;
        EnableQueryRulesLabel: string;
        IncludeOneDriveResultsLabel: string;
        ConnectToSearchVerticalsLabel: string;
        SelectedPropertiesFieldLabel: string;
        SelectedPropertiesFieldDescription: string;
        RefinementFilters: string;
        EnableLocalizationLabel: string;
        EnableLocalizationOnLabel: string;
        EnableLocalizationOffLabel: string;
        QueryCultureLabel: string;
        QueryCultureUseUiLanguageLabel: string;
        InvalidResultSourceIdMessage: string;
        Synonyms: {
            EditSynonymLabel: string;
            SynonymListDescription: string;
            SynonymPropertyPanelFieldLabel: string;
            SynonymListTerm: string;
            SynonymListSynonyms: string;
            SynonymIsTwoWays: string;
            SynonymListSynonymsExemple: string;
            SynonymListTermExemple: string;
        }
    }
}
  
declare module 'SharePointDataSourceStrings' {
  const strings: ISharePointDataSourceStrings;
  export = strings;
}
  