
function to_iso_date(dte) {
	var t = to_js_date(dte);
	return t.toISOString().substr(0,10);
}

function to_xl_date(dte) {
	var t = to_js_date(dte);
	return js_to_excel_date(t);
}

function to_js_date(dte) {
	switch (typeof dte) {
		case "string":
			// from ISO
			return new Date(dte);
		case "number":
			if (dte<99999) {
				// from Excel convention
				return excel_to_js_date(dte);
			} else {
				// from milliseconds since 1-Jan-1970
				return new Date(dte);
			}
		default:
			throw new Error("unhandled date format. "+JSON.stringify(dte));
	}
}

function js_to_excel_date(dte) {
	var t = dte.toUTCString ? dte.valueOf() : dte;
	return Math.floor(t/(24*60*60*1000)) + 25569;
}

function excel_to_js_date(x) {
	var t = (x-25569) * (24*60*60*1000);
	return new Date(t);
}

function remove_time_utc(t) {
	return new Date(t.getFullYear(), t.getMonth(), t.getDate(), -t.getTimezoneOffset()/60);
}

if (typeof exports != "undefined") {
	exports.to_iso_date = to_iso_date;
	exports.to_xl_date = to_xl_date;
	exports.to_js_date = to_js_date;
	exports.remove_time_utc = remove_time_utc;
}