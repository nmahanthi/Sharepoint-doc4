import { app } from '@azure/functions';
import "@pnp/nodejs/index.js";
import "@pnp/sp/webs/index.js";
import "@pnp/sp/items/index.js";
import "@pnp/sp/folders/index.js";
import "@pnp/sp/lists/index.js";

app.setup({
    enableHttpStream: true,
});
