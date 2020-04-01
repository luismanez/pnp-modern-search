export interface IDataSourceData {

    /**
     * Items returned by the data source.
     */
    items: any[];

    /**
     * Any other property returned by the data source to be used in the Handlebars template context.
     */
    [key: string]: any;
}