

function range_to_array(rng) {
	if (Array.isArray(rng) && rng.length>0 && Array.isArray(rng[0])) {
		// that's a 2d array
		var n = rng.length;
		var m = rng[0].length;

		if (n==1) {
			// row array
			return rng[0]
		}
		if (m==1) {
			// column array
			return arr.map(row => row[0]);
		}
		throw new Error("range_to_array: called on "+n+"x"+m+" array. "+JSON.stringify(rng));
	} else if (Array.isArray(rng)) {
		// already a 1d array
		return rng;
	} else {
		// a scalar
		if (rng===null || rng===undefined)
			return rng;
		else
			return [ rng ];
	}
}


function trsp(arr) {
	var n = arr[0].length;
	var m = arr.length;

	var res = new Array(n);
	for (var i=0; i<n; i++) {
		res[i] = new Array(m);
		for (var j=0; j<m; j++)
			res[i][j] = arr[j][i];
	}

	return res;
}

function sort_unique_value(arr) {
	return Array.from(new Set(arr.map(JSON.stringify)).sort().map(JSON.parse);
}

if (typeof exports != "undefined") {
	exports.range_to_array = range_to_array;
	exports.trsp = trsp;
	exports.sort_unique_value = sort_unique_value;
}