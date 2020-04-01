/**
 * Provides useful information about the current Web Part context.
 */
export interface IDataContext {

    /**
     * The input query text if used (ex: from the page environment or search box dynamic data sources)
     */
    inputQueryText: string;

    /**
     * The current selected page number
     */
    pageNumber: number;

    /**
     * The number of items to show per page
     */
    itemsCountPerPage: number;
}