import { app } from '@azure/functions';
import "@pnp/nodejs/index.js";
import "@pnp/sp/sites/index.js";
import "@pnp/sp/webs/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/items/index.js";
import "@pnp/sp/security/index.js";
import "@pnp/sp/fields/index.js";
import "@pnp/sp/site-users/web.js";
import "@pnp/sp/batching.js";

app.setup({
    enableHttpStream: true,
});
