const fs = require("fs");
const path = require("path");

// build with npm install

const addon = require('./build/release/addon');

// add build/Release tot the PATH in order to find xlcall32.dll
console.log("Add build/Release tot the PATH in order to find xlcall32.dll");
//console.log("PATH", process.env.PATH.split(';'));
process.env.PATH = __dirname + "\\build\\Release;" + process.env.PATH;
console.log("PATH", process.env.PATH.split(';'));

console.log("CWD", process.cwd());

/*
required keys:
   name: name of the function as used by Excel and Javascripts
   entry: entry point of the native code function 
   proto: prototye of the function using XLL conventions 
optional keys:
   description: description of the function
   arguments: description 
*/
var global_fct_map = {};

function set_registered_functions(fct_map) {
    for (var f in fct_map) {
        global_fct_map[f] = fct_map[f];
    }
}

function get_registered_functions(xll_path, opt_prms) {
    var prms = opt_prms || {};
    var fct_name = prms.fct_name || "XlXlRegisteredFunctions";
    var fct_type = prms.fct_type || "xloper*";
    var arg_types = prms.arg_types || [ "xloper*" ]; // [ "char*", "xloper*", "double"]
    var arg_vals = prms.arg_vals || []; // [ "abc", [[11,12],[21,22]], 1.234 ]
    
    //console.log("fct_name", fct_name);
    var fcts = addon.xllcall_ffi_v8(xll_path, fct_name, fct_type, arg_types, arg_vals); 
    console.log("#fcts", fcts.length);

    var key_mapping = prms.key_mapping || {
        "procedure": "fct_name",
        "signature": "proto",
        }

    var headers = fcts[0];
    var fct_map = fcts.slice(1).reduce(
        (map,row,i) => {
            var fct_desc = row.reduce( (entry,val,j) => { 
                var key = headers[j].toLowerCase();
                key = key_mapping[key] || key; 
                entry[key] = val; return entry; 
            }, {});
            fct_desc.xll_path = xll_path;
            map[row[0]] = fct_desc;
            return map;
        }, {});
    
    return fct_map;
}

function xllcall_proto_check(xl_name) {
    var fct = global_fct_map[xl_name];
    if (!fct) 
        throw new Error(xl_name+": unknown function (not registered)"); 
    if (fct.fct_type && fct.arg_types) 
        return fct;

    var pos = 0;
    const one_char_pragmas = {
        "!": "!", 
        "#": "#", 
        }; 
    const one_char_types = { 
        P: "xloper*", // array
        R: "xloper*", // ref or array 
        Q: "xloper12*",  // array
        U: "xloper12*",  // ref or array
        B: "double", 
        A: "bool", 
        J: "int32", 
        C: "char*", 
        D: "counted char*",
        };
    const two_char_types = { 
        "C%": "wchar_t*",
        };
    function proto_next_type() {
        var t = two_char_types[fct.proto.substr(pos, 2)];
        if (t) { pos+=2; return t; }
        var t = one_char_types[fct.proto.substr(pos, 1)];
        if (t) { pos++; return t; }
        var t = one_char_pragmas[fct.proto.substr(pos, 1)];
        if (t) { pos++; return; }
        throw new Error(xl_name+": unknown type at pos "+pos+" of prototype string "+fct.proto);
    }

    fct.fct_type = proto_next_type();
    fct.arg_types = [];
    while (pos<fct.proto.length) {
        var next_arg_type = proto_next_type()
        if (next_arg_type)
            fct.arg_types.push(next_arg_type);
    }
    return fct;
}

function xllcall_stub(xl_name) {
    var fct = xllcall_proto_check(xl_name);
    var stub = function() {
        var arg_vals = Array.from(arguments);
        var res = addon.xllcall_ffi_v8(fct.xll_path, fct.fct_name, fct.fct_type, fct.arg_types, arg_vals); 
        return res;
    }
    return stub;
}

function xldna_test(xldna_path) {
    set_registered_functions({
        RegistrationInfo : { fct_name: "RegistrationInfo", proto: "QQ", xll_path: xldna_path },
        SyncMacro : { fct_name: "SyncMacro", proto: ">B", xll_path: xldna_path },
        f0 : { fct_name: "f0", proto: "PCB", xll_path: xldna_path },
        f1 : { fct_name: "f1", proto: "QC%B", xll_path: xldna_path },
    });

    var RegistrationInfo =  xllcall_stub("RegistrationInfo");
    //var SyncMacro =  xllcall_stub("SyncMacro");
    var f0_addthem_97 =  xllcall_stub("f0");

    //var res = RegistrationInfo();
    //console.log("RegistrationInfo", res);
    
    var res = f0_addthem_97("sa",3);
    console.log("addthem", res);    
}
// test 
if (!module.parent) {
    // module.parent is null, we're running the main module

    if (1) {
        var xldna_path = path.join(__dirname, "ExcelDna\\ExcelDna-0.34.6\\ExcelDna\\Distribution\\ExcelDna64.xll");
        console.log("xldna_path", xldna_path, fs.existsSync(xldna_path));
        
        var xldna_folder = xldna_path.split("\\"); xldna_folder.pop();
        xldna_folder = xldna_folder.join("\\");
        console.log("xldna_folder", xldna_folder, fs.existsSync(xldna_folder));
        process.env.PATH = xldna_folder + ";" + process.env.PATH;
        
        var xldna_res = xldna_test(xldna_path);
        console.log("xldna_res", xldna_res);
    }

    if (0) {
        var xll_path = path.join(__dirname, "Vision.xll");
        console.log("xll_path", xll_path, fs.existsSync(xll_path));
        
        var fct_map = get_registered_functions(xll_path);
        console.log("fct_map",  JSON.stringify(Object.keys(fct_map), null, 4));

        set_registered_functions(fct_map);
        for (f in fct_map)
            xllcall_stub(f);

        var XlSet = xllcall_stub("XlSet");
        var XlGet = xllcall_stub("XlGet");
        
        console.log("XlSet1", XlGet(XlSet("tkr", "abc")));
        console.log("XlSet2", XlGet(XlSet("tkr", 123)));
        console.log("XlSet3", XlGet(XlSet("tkr", true)));
        //try { console.log("XlSet4", XlGet(XlSet("tkr", [ 1, 2, 3 ]))); } catch(e) { console.warn("XlSet4", e)}
        console.log("XlSet5", XlGet(XlSet("tkr", [[ 11, 12 ], [21, 22 ]])));
        console.log("XlSet6", XlGet(XlSet([["tkr"]], [[ 11, 12 ], [21, 22 ]])));
    }
}

//exports.xllcall_ffi = xllcall_ffi_v8; // invoke the native implementation
exports.xllcall_ffi = function() {
    return addon.xllcall_ffi_v8(); // invoke the native implementation
}