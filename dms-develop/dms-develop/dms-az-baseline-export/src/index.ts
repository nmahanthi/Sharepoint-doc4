import { app } from '@azure/functions';
import "az-common/node_modules/@pnp/sp/presets/all.js";
import "az-common/node_modules/@pnp/sp/batching.js";
import "az-common/node_modules/@pnp/graph/presets/all.js";
import "az-common/node_modules/@pnp/graph/batching.js";

app.setup({
    enableHttpStream: true,
});
