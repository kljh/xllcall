'use strict';  

const fs = require("fs");
const path = require("path");

// build with npm install

const addon = require('./build/release/addon');

// set whether the XLLCall should be verbose (C++)
addon.xllcall_debug_v8(false); 

// add build/Release tot the PATH in order to find xlcall32.dll
//console.log("Add build/Release tot the PATH in order to find xlcall32.dll");
process.env.PATH = __dirname + "\\build\\Release;" + process.env.PATH;
//console.log("PATH", process.env.PATH.split(';'));

//console.log("Current Working Directory", process.cwd());

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
    var arg_names = prms.arg_names;
    var arg_vals = prms.arg_vals || []; // [ "abc", [[11,12],[21,22]], 1.234 ]
    
    //console.log("fct_name", fct_name);
    var fcts = addon.xllcall_ffi_v8(xll_path, fct_name, fct_type, arg_types, arg_names, arg_vals); 
    console.log("#fcts", fcts.length);

    var key_mapping = prms.key_mapping || {
        "procedure": "fct_name",
        "signature": "proto",
        "module": "xll_path",
        "entry": "fct_name",
        }

    var headers = fcts[0];
    var fct_map = fcts.slice(1).reduce(
        (map,row,i) => {
            var fct_desc = row.reduce( (entry,val,j) => { 
                var key = headers[j].toLowerCase();
                key = key_mapping[key] || key; 
                entry[key] = val; return entry; 
            }, {});
            fct_desc.xll_path = fct_desc.xll_path || xll_path;
            map[row[0]] = fct_desc;
            return map;
        }, {});
    
    return fct_map;
}

function xllcall_proto_check(xl_name) {
    var fct = global_fct_map[xl_name];
    if (!fct) {
        //console.warn("Registered functions:", Object.keys(global_fct_map).sort());
        throw new Error(xl_name+": unknown function (not registered)"); 
    }
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
        throw new Error(xl_name+": unknown type "+fct.proto.substr(pos, 3)+"... at pos "+pos+" of prototype string "+fct.proto);
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
        try {
            var res = addon.xllcall_ffi_v8(fct.xll_path, fct.fct_name, fct.fct_type, fct.arg_types, fct.arg_names, arg_vals); 
        } catch (e) {
            var nb_args = arg_vals.length;
            console.error("Error while calling "+xl_name, fct, !nb_args ? " with no args." : " with args:");
            for (var a=0; a<nb_args; a++)
                console.error("arg "+(a+1)+"/"+nb_args+" "+(fct.arg_names?fct.arg_names[a]:undefined)+": ", arg_vals[a]);
            throw e;
        }
        return res;
    }
    return stub;
}

function xldna_init(xldna_xll_path) {
    console.log("process.arch", process.arch);
    var path_suffix = process.arch=="x64" ? "64" : "";
    var xldna_path = xldna_xll_path || path.join(__dirname, "ExcelDna\\ExcelDna-0.34.6\\ExcelDna\\Distribution\\ExcelDna"+path_suffix+".xll");
    var xldna_path = xldna_xll_path || path.join(__dirname, "ExcelDna\\ExcelDna-0.34.6\\ExcelDna\\Source\\ExcelDna\\Debug"+path_suffix+"\\ExcelDna"+path_suffix+".xll");
    if (!fs.existsSync(xldna_path)) 
        throw new Error("xldna_path does not exist. "+xldna_path);
    
    var xldna_folder = path.dirname(xldna_path);
    process.env.PATH = xldna_folder + ";" + process.env.PATH;
    
    addon.xllload_xlAutoOpen_v8(xldna_path);

    set_registered_functions({
        RegistrationInfo : { fct_name: "RegistrationInfo", proto: "QQ", xll_path: xldna_path },
        SyncMacro : { fct_name: "SyncMacro", proto: ">B", xll_path: xldna_path },
        f0 : { fct_name: "f0", proto: "PCB", xll_path: xldna_path },
        f1 : { fct_name: "f1", proto: "QC%B", xll_path: xldna_path },
    });

    // run test
    if (1) {
        var RegistrationInfo =  xllcall_stub("RegistrationInfo");
        //var SyncMacro =  xllcall_stub("SyncMacro");
        var f0_addthem_97 =  xllcall_stub("f0");

        var res = RegistrationInfo();
        if (res) res = res.map(row => row.slice(0,16));
        console.log("RegistrationInfo (partial rows):", res);
        
        var res = f0_addthem_97("Lupin dog no", 1);
        console.log("addthem", res);
    };
}

function vision_init(vision_xll_path) {
    var xll_path = vision_xll_path || path.join(__dirname, process.arch=="x64" ? "bin\\x64\\vision.xll" : "bin\\x86\\vision.xll"); // process.env.APPDATA
    if (!fs.existsSync(xll_path)) 
        throw new Error("xll_path does not exist. "+xll_path);
    
    var fct_map = get_registered_functions(xll_path);
    set_registered_functions(fct_map);
    // console.log("fct_map",  JSON.stringify(Object.keys(fct_map), null, 4));
    console.log("#fcts",  Object.keys(fct_map).length);
    
    // force stub creation of all functions, and define them ing lobal scope
    if (1) {
        for (var f in fct_map) 
            global[f] = xllcall_stub(f);
    }

    // run test
    if (1) {
        var XlSet = xllcall_stub("XlSet");
        var XlGet = xllcall_stub("XlGet");
        var XlAutoFit = xllcall_stub("XlAutoFit");
        
        //console.log(fct_map["XlSet"])
        console.log("XlSet1", XlGet(XlSet("tkr", "abc")));
        console.log("XlSet2", XlGet(XlSet("tkr", 123)));
        console.log("XlSet3", XlGet(XlSet("tkr", true)));
        //try { console.log("XlSet4", XlGet(XlSet("tkr", [ 1, 2, 3 ]))); } catch(e) { console.warn("XlSet4", e)}
        console.log("XlSet5", XlGet(XlSet("tkr", [[ 11, 12 ], [21, 22 ]])));
        console.log("XlSet6", XlGet(XlSet([["tkr"]], [[ 11, 12 ], [21, 22 ]])));

        // this function calls back Excel to ask the size of the target range
        console.log("XlAutoFit", XlAutoFit("abc"));
        console.log("XlAutoFit", XlAutoFit([[ "abc", 123 ], [ true, 456 ]]));
    }
}

function top_left(x) {
    if (Array.isArray(x)) {
        if (x.length > 0)
            return top_left(x[0]);
        else
            return;
    } else {
        return x;
    }
}

// test 
if (!module.parent) {
    xldna_init();
    //vision_init();
}

exports.xllcall_debug = addon.xllcall_debug_v8;   // invoke the native implementation
exports.xllcall_ffi = addon.xllcall_ffi_v8;       // invoke the native implementation
exports.xllcall_stub = xllcall_stub;
exports.xllcall_proto_check = xllcall_proto_check;
exports.global_fct_map = global_fct_map;
exports.get_registered_functions = get_registered_functions;
exports.set_registered_functions = set_registered_functions;
exports.vision_init = vision_init;
exports.xldna_init = xldna_init;
exports.top_left = top_left;
