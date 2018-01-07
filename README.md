# xllcall

[![NPM Version][npm-image]][npm-url]
[![NPM Downloads][downloads-image]][downloads-url]

Invoke XLL functions from Node.

## Installation

```sh
npm install xllcall
```

or, to force 32 bit build for x86 XLLs (requires 32 bit Node)

```sh
npm install xllcall --arch=ia32
```

## Running from Visual Studio Code (to be able to debug) 

Run server.js, that will start a HTTP server that can receive requests.

If running a 64 bit system and intending to use 32 bit XLLs, then you must modify you launch.json.
Specify a 32 bit version of node.exe in runtimeExecutable :
```sh
    {
        "version": "0.2.0",
        "configurations": [
            {
                "type": "node",
                "request": "launch",
                "runtimeExecutable": "C:\\node-v8.9.4-win-x86\\node.exe",
                "name": "Launch Program",
                "program": "${file}"
            }
        ]
    }
```

## API

```js
var xllcall = require('xllcall');

// load a XLL triggers the registration of the XLL functions.. 
xllcall.xllload("myxll.xll");
// the XLL functions can then be called as with Application.Run
xllcall.xllcall("myfct", "abc", [[11,12],[21,22]], 1.234 );

// the list of registered functions is available 
var fcts = xllcall.xllfcts()

// one can also call a function explicitly specifying the signature
xllcall.xllcall_ffi("myxll.xll", "myfct", [ "xloper", "const char*", "xloper", "double"], [ "abc", [[11 12],[21,22]], 1.234 ]);
```

### xllload(xll_name)

### xllcall(fct_name, args_values, ...)

### xllfcts()

### xllcall_ffi(xll_name, fct_name, [ return_type, arg_types, ...], [ arg_values, ... ])

### Conversion Rules between JS and XLL

Argument conversion depends on the expected type for the argument:

* string arguments:
  * undefined and null become empty strings
  * numbers and booleans become their text representation
  * string are assumed to only contain latin1 character (XLL using XL97 API) or UCS2 characters (XLL using XL2007 API)
  * array: first value is used (top left value for 2D array)

* number arguments:
  * undefined and null become empty strings
  * boolean : true become 1, and false becomes 0
  * array: first value is used (top left value for 2D array)

* Boolean arguments
  * undefined and null become false
  * non-zero numbers are become true, zeros become false

* Range arguments:
  * one dimensionnal array becomes column ranges



[npm-image]: https://img.shields.io/npm/v/xllcall.svg
[npm-url]: https://npmjs.org/package/xllcall
[downloads-image]: https://img.shields.io/npm/dm/xllcall.svg
[downloads-url]: https://npmjs.org/package/xllcall
