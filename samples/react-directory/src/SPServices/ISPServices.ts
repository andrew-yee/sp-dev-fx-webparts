// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { PeoplePickerEntity } from '@pnp/sp';

export interface ISPServices {

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    searchUsers(searchString: string, searchFirstName: boolean): any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    searchUsersNew(searchString: string, srchQry: string, isInitialSearch: boolean, pageNumber?: number): any;

}
