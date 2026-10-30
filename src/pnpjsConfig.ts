
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ExtensionContext } from '@microsoft/sp-extension-base';

let _sp: SPFI | undefined;

export const getSP = (context?: ExtensionContext): SPFI => {
    if (context) {
        _sp = spfi().using(SPFx(context));
    }

    if (!_sp) {
        throw new Error('PnPjs has not been initialized.');
    }

    return _sp;
};
