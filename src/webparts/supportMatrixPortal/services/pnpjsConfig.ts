import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import { WebPartContext } from '@microsoft/sp-webpart-base';

let _sp: SPFI;

/**
 * Initializes the shared PnPjs SPFI instance. Must be called once from the web part's
 * onInit() before any service call is made.
 */
export const initializePnPjs = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};

/**
 * Returns the shared PnPjs SPFI instance.
 */
export const getSP = (): SPFI => {
  if (!_sp) {
    throw new Error('PnPjs has not been initialized. Call initializePnPjs(context) first.');
  }
  return _sp;
};
