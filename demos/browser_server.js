/**
 * local web server
 * since we moved to modules (`jsm`) with v3.6.0, merely opening the local file in browser gices CORS errors
 */
const express = require("express");
const app = express();
const port = 8000;

app.use("/browser", express.static("./browser"));
app.use("/common", express.static("./common"));
app.use("/modules", express.static("./modules"));

app.listen(port, () => {
	console.log(`[local server] listening on port ${port}!`);
});
