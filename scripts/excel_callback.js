const request = require('request');
const rpm = require('request-promise');

const xl_callback_url = "http://localhost:9707/"

function help() {
	var url = xl_callback_url + "?cmd=help&sequence_id=" + this.req.query.sequence_id
	return rpm(url)
	.catch(request_promise_reject_handler);
}

function list() {
	var url = xl_callback_url + "?cmd=list&sequence_id=" + this.req.query.sequence_id
	return rpm(url)
	.catch(request_promise_reject_handler);
}

function xlget(id) {
	var url = xl_callback_url + "?cmd=xl_get&sequence_id=" + this.req.query.sequence_id + "&id=" +  id
	return rpm(url)
	.catch(request_promise_reject_handler);
}

function xludf(udf_name, args, opt_timeout) {
	if (!args) args = [];
	console.log("xludf "+JSON.stringify(udf_name)+"( "+args.map(JSON.stringify).join(", ")+" ) ...");

	var url = xl_callback_url + "xludf?sequence_id=" + this.req.query.sequence_id;
	var body = { "udf_name": udf_name };
	for (var i=0; i<args.length; i++)
		body["arg"+i] = args[i];

	console.log("calls back Xl with URL "+url+"...");

	return rpm({
		method: "POST",
		uri: url,
		body: body,
		json: true,
		timeout: opt_timeout || 5000, // 5 s.
		//resolveWithFullResponse: true
	})
	.then(response => (response && response.result) || response)
	.catch(request_promise_reject_handler);
}

function request_promise_reject_handler(err) {
	var { name, statusCode, body, message, error, } = err; // extract
	console.log("xlcbk ERROR "+JSON.stringify({ name, statusCode, body, message, error, }, null, 4));
	return Promise.reject({ name, statusCode, body, message, error, });
}

async function test() {
	console.log();

	console.log("help");
	console.log(await help.apply(this));
	console.log();

	console.log("list");
	console.log(await help.call(this));
	console.log();

	console.log("xludf XlToday()");
	console.log(await xludf.call(this, "XlToday"));
	console.log();

	console.log("xludf XlToday(47000)");
	console.log(await xludf.call(this, "XlToday", [4700]));
	console.log();

	console.log("xludf XlAggRow('abc', true, [[11, 12],[21,22]])");
	console.log(await xludf.call(this, "XlAggRow", [ 'abc', true, [[11, 12],[21,22]], null ]));
	console.log();

	return "OK"
}

if (typeof exports != undefined) {
	exports.help = help;
	exports.xludf = xludf;
	exports.test = test;

}
