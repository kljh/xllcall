'use strict';

const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const url = require('url');
const express = require('express');
const session = require('express-session')
const bodyParser = require('body-parser');

// path to scripts
const script_root = path.join(__dirname, "scripts");
const script_prefix = "";
const script_suffix = ".js"

// addon to call XLL
const xll_util = require('./index');
xll_util.vision_init();

// addon to run a webserver
const app = express();
const server = http.createServer(app);

const http_port = 8664;
server.listen(http_port, function () { console.log('HTTP server started on port: %s', http_port); });

app.use(bodyParser.text())
app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json())

app.use(function (req, res, next) {
    switch (req.path) {
        case "/xll_registered_functions":
            var fcts = xll_util.global_fct_map;
            for (var fct in fcts) xll_util.xllcall_proto_check(fct);
            return  res.jsonp(fcts);
        case "/xll_invoke_function":
            var fct_name = req.query.fct;
            var fct = xll_util.global[fct_name];
            var args = req.query.args;

            //var ret = fct.apply(null, args);
            //return res.jsonp(ret);

            Promise
            .resolve(fct.apply(null, args))
            .then(ret =>res.jsonp(ret))
            .catch(next);

            break;
        default:
            next();
    }
});

app.use(function (req, res, next) {
    try {
        //console.log("request received.")
        var request = req.body.substr ? JSON.parse(req.body) : req.body;
        //console.log("request body.", request);
        console.log("request received", request.request_name);

        if (!request.request_name) throw new Error("request_name missing.");
        var request_src = request.request_name.split(".").shift();
        var request_fct = request.request_name.split(".").pop();

        var src_file = path.join(script_root, script_prefix+request_src+script_suffix);

        /*
        // Poor man's style modules
        var src_text = fs.readFileSync(src_file, "utf8");
        //console.log("src_text.", src_text);
        var fct = new Function("request", src_text+"return "+request_fct+"(request);");
        var response = fct(request);
        */

        // Node style modules
        //delete require.cache[require.resolve(src_file)]; // force reload
        for (var script_path in require.cache) {
            if (script_path.indexOf(script_root)==0) {
                console.log("clearing "+script_path+" from cache");
                delete require.cache[script_path];
            }
        }


        var that = {req};
        var src = require(src_file);
        var fct = src[request_fct];
        if (!fct) throw new Error("function "+request_fct+" not exported by script "+request_src+".");
        var response;
        if (request.ordered_args) {
            var response = fct.apply(that, request.ordered_args);
        } else if (request.named_args) {
            var function_arg_names = (""+fct).match(/function\s*[^\(]*\([^\)]*\)/)[0].split(",").map(arg => arg.trim());
            var ordered_args = function_arg_names.map(arg_name => request.named_args[arg_name]);
            var response = fct.apply(that, ordered_args);
        } else {
            var response = fct.apply(that, [ request ]);
        }

        Promise
        .resolve(response)
        .then(response_resolved => {
            var response_json;
            switch (typeof response_resolved) {
                case "object" :
                    response_json = JSON.stringify(response_resolved,null,2); break;
                case "undefined":
                    response_json = "null";
                default:
                    response_json = "" + response_resolved;
            }
            res.send(response_json);
            console.log("response sent. #"+response_json.length+"\n")
        })
        .catch(next);

    } catch (e) {
        console.error("request error sent. "+e+".\n", request, "\n");
        res.status(500).send(JSON.stringify({
                error: ""+e,
                stack: e.stack.split('\n'),
            }, null, 4));
    }
});
