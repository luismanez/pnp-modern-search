import { IDataSource } from "./IDataSource";
import { IDataSourceData } from "./IDataSourceData";
import { IPropertyPaneGroup, IPropertyPaneConditionalGroup } from "@microsoft/sp-property-pane";
import { ServiceScope } from '@microsoft/sp-core-library';
import { IDataContext } from './IDataContext';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export abstract class BaseDataSource<T> implements IDataSource {

    protected serviceScope: ServiceScope;

    protected _properties: T;
    private _context: WebPartContext;
    private _render: () => void | Promise<void>;

    get properties(): T {
        return this._properties;
    }

    set properties(properties: T) {
        this._properties = properties;
    }

    get render(): () => void | Promise<void> {
        return this._render;
    }

    set render(renderFunc: () => void | Promise<void>) {
        this._render = renderFunc;
    }

    get context(): WebPartContext {
        return this._context;
    }

    set context(context: WebPartContext) {
        this._context = context;
    }

    public constructor(serviceScope: ServiceScope) {
        this.serviceScope = serviceScope;
    }

    public onInit(): void | Promise<void> {
    }

    public abstract async getData(dataContext?: IDataContext): Promise<IDataSourceData>;

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] | IPropertyPaneConditionalGroup[] {

        // Returns an empty configuration by default
        return [];
    }

    public abstract getItemCount(): number;

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        // Do nothing by default      
    }

    public onPropertyPaneConfigurationStart(): Promise<void> {
        return;
    }
}