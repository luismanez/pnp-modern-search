import { IDataSourceData } from "./IDataSourceData";
import { IPropertyPaneGroup } from "@microsoft/sp-property-pane";
import { IDataContext } from './IDataContext';
import { ServiceKey } from "@microsoft/sp-core-library";
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDataSource {

    /**
     * The Web Part properties in the property bag. Corresponds to the isolated 'dataSourceProperties' property in the global property bag.
     */
    properties: any;

    /**
     * This API is called to render the web part.
     */
    render: () => void | Promise<void>;

    /**
     * Context of the main Web Part
     */
    context: WebPartContext;

    /**
     * Method called during the Web Part initialization.
     */
    onInit(): void | Promise<void>;

    /**
     * Retrieves the data from this data source.
     * @param dataContext useful information about the current Web Part context (for instance, current page number, etc.).
     */
    getData(dataContext?: IDataContext): Promise<IDataSourceData>;

    /**
     * Returns the total of items.
     */
    getItemCount(): number;
    
    /**
     * Method called when a property pane field in changed in the Web Part.
     * @param propertyPath the property path.
     * @param oldValue the old value.
     * @param newValue the new value.
     */
    onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void;

    /**
     * Method called when the property pane configuration starts
     */
    onPropertyPaneConfigurationStart(): Promise<void>;
}