/**
 * NAME: browser_server.mjs
 * DESC: Local web server for ./demos/browser/index.html
 * DESC: since we moved to modules (`jsm`) with v3.6.0, merely opening the local file in browser gives CORS errors
 * REQS: express
 * USAGE: `node browser_server.mjs`
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20210404
 */

import express from "express";
const app = express();
const port = 8000;

app.use("/browser", express.static("./browser"));
app.use("/common", express.static("./common"));
app.use("/modules", express.static("./modules"));

app.listen(port, () => {
	console.log(`\n----------------------==~==~==~==[ SERVER RUNNING ]==~==~==~==----------------------\n`);
	console.log(`The pptxgenjs browser demo is now live at: http://localhost:${port}/browser/index.html\n`);
	console.log(`(Press Ctrl-C to stop)`);
	console.log(`\n----------------------==~==~==~==[ SERVER RUNNING ]==~==~==~==----------------------\n`);
});
