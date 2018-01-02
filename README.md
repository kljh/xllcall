# xllcall

[![NPM Version][npm-image]][npm-url]
[![NPM Downloads][downloads-image]][downloads-url]

Invoke XLL functions from Node.

## Installation

```sh
npm install xllcall
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


[npm-image]: https://img.shields.io/npm/v/xllcall.svg
[npm-url]: https://npmjs.org/package/xllcall
[downloads-image]: https://img.shields.io/npm/dm/xllcall.svg
[downloads-url]: https://npmjs.org/package/xllcall
