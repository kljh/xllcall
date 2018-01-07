'use strict';  

const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const url = require('url');
const express = require('express');
const session = require('express-session')
const bodyParser = require('body-parser');

// path to XLL and to scripts
const script_root = path.join(__dirname, "..");
const script_prefix = "";
const script_suffix = ".js"

// addon to call XLL
const xll_util = require('./index');
xll_util.xll_init();

// addon to run a webserver
const app = express();
const server = http.createServer(app);

const http_port = 8664;
server.listen(http_port, function () { console.log('HTTP server started on port: %s', http_port); });

app.use(bodyParser.text())
app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json())

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
        delete require.cache[require.resolve(src_file)]; // force reload
        var src = require(src_file);
        var fct = src[request_fct];
        if (!fct) throw new Error("function "+request_fct+" not exported by script "+request_src+".");
        var response = fct(request);
        
        var response_json = JSON.stringify(response,null,2);
        res.send(response_json);
        console.log("response sent. #"+response_json.length+"\n")
    } catch (e) {
        console.error("request error sent. "+e+".\n", request, "\n");
        res.status(500).send(JSON.stringify({
                error: ""+e, 
                stack: e.stack.split('\n'),
            }, null, 4));
    }
});



