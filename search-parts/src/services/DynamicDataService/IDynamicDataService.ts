import IDataSourceProperty from "../../models/IDataSourceProperty";
import { DynamicDataProvider } from "@microsoft/sp-component-base";

export default interface IDynamicDataService {
    dynamicDataProvider: DynamicDataProvider;
    getAvailableDataSourcesByType(propertyId: string): IDataSourceProperty[];
}