var https       = require('https');
var url         = require('url');
var xpath       = require('xpath');
var cookie      = require('cookie');
var Cache       = require("mem-cache");
var uuid        = require("node-uuid");
var domParser   = new (require('xmldom').DOMParser)();

var Util = function (settings) {

    if (!settings || typeof(settings)!=="object") throw new Error("'settings' argument must be an object instance.");
    if (!settings.host || typeof(settings.host)!=="string") throw new Error("'settings.host' property is a required string.");
    if (settings.timeout!==undefined && typeof(settings.timeout)!=="number") throw new Error("'settings.timeout' property must be a number.");

    var self = this;
    var loginPath       = '/_forms/default.aspx?wa=wsignin1.0';
    var loginEndpoint   = url.resolve("https://" + settings.host, loginPath);
    var entitySets      = null;
    var pendingHook     = null;

    settings.timeout = settings.timeout || 15 * 60 * 1000;     // default sessions timeout of 15 minutes in ms   

    Object.defineProperty(this, "cacheAuth", {
        enumerable: false,
        configurable: false,
        writable: false,
        value: new Cache(settings.timeout)
    });

    var cacheUser   = new Cache(settings.timeout);   // Cache by auth tokens 

    this.samlTemplate = '\
        <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">\
            <s:Header>\
                <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>\
                <a:ReplyTo>\
                    <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>\
                </a:ReplyTo>\
                <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>\
                <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">\
                    <o:UsernameToken>\
                        <o:Username>{username}</o:Username>\
                        <o:Password>{password}</o:Password>\
                    </o:UsernameToken>\
                </o:Security>\
            </s:Header>\
            <s:Body>\
                <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">\
                  <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">\
                    <a:EndpointReference>\
                      <a:Address>{endpoint}</a:Address>\
                    </a:EndpointReference>\
                  </wsp:AppliesTo>\
                  <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>\
                  <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>\
                  <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>\
                </t:RequestSecurityToken>\
            </s:Body>\
        </s:Envelope>';

    this.authenticate = function(options, cb) {

        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        cb = cb || defaultCb;
        options = options || {};
        if (typeof options !== 'object') return cb(new Error("'options' argument is missing or invalid."));

        // Validates username and password 
        options.username = options.username || settings.username;
        options.password = options.password || settings.password;

        if (!(options.username)) return cb( new Error("'options.username' property is required."));
        if (!(options.password)) return cb( new Error("'options.password' property is required."));

        var auth = cacheUser.get(options.username);
        if (auth) {
            var item = cacheAuth.get(auth);
            if (item && item.password === options.password) return cb(null, item.user);
        }

        getAuthenticationToken(options.username, options.password, function (err, authToken) {

            if (err) {
                // cleans caches
                if (auth) {
                    cacheAuth.remove(auth);
                    cacheUser.remove(options.username)
                }
                return cb(err);
            }

            getAuthorization(authToken, function (err, authz) {

                if (err) {
                    // cleans caches
                    if (auth) {
                        cacheAuth.remove(auth);
                        cacheUser.remove(options.username)
                    }
                    return cb(err);
                }

                // reuse auth by username
                auth = auth || uuid.v4();

                // creates cache item
                var item = {
                    authz       : authz,
                    cookieAuthz : 'FedAuth=' + authz.FedAuth + ';rtFa=' + authz.rtFa,
                    username    : options.username,
                    password    : options.password
                };

                self.cacheAuth.set(auth, item);
                cacheUser.set(options.username, auth);
                
                getEntitySets(auth, function(err) {
                    if (err) return cb(err);
                    cb(null, { auth: auth });
                });
            });
        });
    };


    this.oData = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options || typeof options !== 'object')                    return cb(new Error("'options' argument is missing or invalid."));
        if (!options.command || typeof options.command !== 'string')    return cb(new Error("'options.command' argument is missing or invalid."));

        request(options, cb);
    };


    this.entitySets = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options || typeof options !== 'object')    return cb(new Error("'options' argument is missing or invalid."));

        if (options.auth) {

            getEntitySets(options.auth, cb);
        } else {

            this.authenticate(options, function(err, result) {
                if (err) return cb(new Error("Couldn't get the entity sets. " + err));
                getEntitySets(result.auth, cb);
            });
        }
    };


    this.get = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource!=='string') return cb(new Error("'resource' property is missing or invalid."));
        if (options.id === null || options.id === undefined) return cb(new Error("'id' property is missing."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            method  : "GET",
            command : "/" + options.resource + "(" + buildODataId(options.id) + ")"
        }, cb);
    };


    this.query = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource!=='string') return cb(new Error("'resource' property is missing or invalid."));

        var err = null;
        var params = ["filter", "expand", "select", "orderBy", "top", "skip"]
            .map(function (prop) {
                if (options[prop]) {
                    var value = options[prop];
                    if (typeof value !== 'string') {
                        err = new Error("The property '" + prop + "' must be a valid string.");
                        return null;
                    }
                    return "$" + prop.toLowerCase() + "=" + value;
                }
                return null;
            })
            .filter(function (param) { return !!param; });

        if (err) return cb(err);

        params.push("$inlinecount=" + (options.inLineCount ? "allpages" : "none"));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            method  : "GET",
            command : "/" + options.resource + (params.length === 0 ? "" : "?" + params.join("&"))
        }, cb);
    };


    this.links = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined)            return cb(new Error("'id' property is missing."));
        if (!options.resource || typeof options.resource !== 'string')  return cb(new Error("'resource' property is missing or invalid."));
        if (!options.entity || typeof options.entity !== 'string')      return cb(new Error("'entity' property is missing or invalid."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            method  : "GET",
            command : "/" + options.resource + "(" + buildODataId(options.id) + ")/$links/" + options.entity
        }, cb);
    };


    this.count = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource !== 'string')  return cb(new Error("'resource' property is missing or invalid."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            method  : "GET",
            command : "/" + options.resource + "/$count"
        }, cb);
    };


    this.create = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (options.data === null || options.data === undefined)     return cb(new Error("'data' property is missing."));
        if (!options.resource || typeof options.resource!=='string') return cb(new Error("'resource' property is missing or invalid."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            data    : options.data,
            method  : "POST",
            command : "/" + options.resource
        }, cb);
    };


    this.replace = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined)            return cb(new Error("'id' property is missing."));
        if (options.data === null || options.data === undefined)        return cb(new Error("'data' property is missing."));
        if (!options.resource || typeof options.resource!=='string')    return cb(new Error("'resource' property is missing or invalid."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            data    : options.data,
            method  : "PUT",
            command : "/" + options.resource + "(" + buildODataId(options.id) + ")"
        }, cb);
    };


    this.update = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined)            return cb(new Error("'id' property is missing."));
        if (options.data === null || options.data === undefined)        return cb(new Error("'data' property is missing."));
        if (!options.resource || typeof options.resource!=='string')    return cb(new Error("'resource' property is missing or invalid."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            data    : options.data,
            method  : "PATCH",
            command : "/" + options.resource + "(" + buildODataId(options.id) + ")"
        }, cb);
    };


    this.remove = function(options, cb) {
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }
        
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource!=='string')    return cb(new Error("'resource' property is missing or invalid."));
        if (options.id === null || options.id === undefined)            return cb(new Error("'id' property is missing."));

        self.oData({
            auth    : options.auth,
            username: options.username,
            password: options.password,
            method  : "DELETE",
            command : "/" + options.resource + "(" + buildODataId(options.id) + ")"
        }, cb);
    };


    this.hook = function(target) {

        var addMethod = function(method, prefix, entitySet) {

            target[prefix + entitySet] = function (options, cb) {

                if (!cb && typeof options === 'function') {
                    cb = options;
                    options = {};
                }
                
                cb = cb || defaultCb;
                options = options || {};
                options.resource = entitySet;

                method(options, cb);
            };
        };

        var addAllMethods = function(entitySets) {
            entitySets.forEach( function (entitySet) {
                addMethod(self.get, "get", entitySet);
                addMethod(self.query, "query", entitySet);
                addMethod(self.links, "links", entitySet);
                addMethod(self.count, "count", entitySet);
                addMethod(self.create, "create", entitySet);
                addMethod(self.replace, "replace", entitySet);
                addMethod(self.update, "update", entitySet);
                addMethod(self.remove, "remove", entitySet);
            });
        };

        if (entitySets) {
            addAllMethods(entitySets);
        } else {
            pendingHook = function() {
                addAllMethods(entitySets);
                pendingHook = null;
            }
        }
    };


    var buildODataId = function(id) {

        if (typeof id === 'string') return "'" + id + "'";
        return "" + id;        

    }
    var getEntitySets = function(auth, cb) {

        if (entitySets) return cb(null, entitySets);

        self.oData({ auth: auth, command: "$metadata"}, function(err, result) {

            try {

                if (err) throw err;

                entitySets = xpath
                    .select("//*[local-name(.)='EntitySet' and namespace-uri(.)='http://schemas.microsoft.com/ado/2009/11/edm']/@Name", result.data)
                    .map(function(attr) { return attr.value;} );

                if (pendingHook) pendingHook();

                cb(null, entitySets);

            } catch(e) {

                entitySets = null;
                cb(new Error("Couldn't get entity sets. " + e));
            }
        });
    
    }

    var name = function(name) {
        return "/*[name(.)='" + name + "']";
    };

    var getAuthenticationToken = function(username, password, cb) {

        var samlRequest = self.samlTemplate
            .replace("{username}", username)
            .replace("{password}", password)
            .replace("{endpoint}", loginEndpoint);

        var options = {
            method: 'POST',
            host: 'login.microsoftonline.com',
            path: '/extSTS.srf',
            headers: { 'Content-Length': samlRequest.length }
        };

        var req = https.request(options, function (res) {

            var xml = '';
            res.setEncoding('utf8');
            res.on('data', function (chunk) { xml += chunk; })
            res.on('end', function () {

                var resXml = domParser.parseFromString(xml); 
                var exp = ['S:Envelope', 'S:Body', 'S:Fault', 'S:Detail', 'psf:error', 'psf:internalerror', 'psf:text'].map(name).join("") + "/text()";
                var fault = xpath.select(exp, resXml);
                if (fault.length > 0) return cb(new Error(fault.toString()));
        
                exp = ['S:Envelope', 'S:Body', 'wst:RequestSecurityTokenResponse', 'wst:RequestedSecurityToken', 'wsse:BinarySecurityToken'].map(name).join("") + "/text()";
                var token = xpath.select(exp, resXml);
                if (token.length > 0) return cb(null, token.toString());
            })
        });
        
        req.end(samlRequest);
    };

    var getAuthorization = function(authToken, cb) {

        var options = {
            method: 'POST',
            host: settings.host,
            path: loginPath,
            headers: {
                'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0)',
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        };

        var req = https.request(options, function (res) {

            res.setEncoding('utf8');
            res.on('end', function () {
                var cookies = cookie.parse(res.headers["set-cookie"].join(";"));
                cb (null, {
                    FedAuth : cookies.FedAuth,
                    rtFa    : cookies.rtFa                    
                });
            });
        });

        req.end(authToken);
    };

    var buildPath = function(command) {
        var path = (settings.site ? "/" + settings.site : "") + "/_api/" + command;

        while (path.indexOf("//") >- 1) {
            path = path.replace("//", "/");
        }

        return encodeURI(path);
    };

    var request = function(options, cb) {

        if (!options.auth) {

            self.authenticate({ username: options.username, password: options.password }, function (err, result) {
                if (err) return cb(err);

                // tries again
                options.auth = result.auth;
                request(options, cb);
            });

        } else {

            // gets cached authorization data
            var item = self.cacheAuth.get(options.auth);
            if (!item) return cb (new Error("Invalid 'auth' property."));

            var data = options.data ? JSON.stringify(options.data) : "";

            var reqOptions = {
                method  : options.method || "GET",
                host    : settings.host,
                path    : buildPath(options.command),
                headers : {
                    'accept'        : 'application/json;odata=verbose',
                    'cookie'        : item.cookieAuthz,
                    'content-type'  : 'application/json',
                    'content-length': data.length
                }
            };

            if (options.etag) {
                reqOptions.headers['if-match'] = options.etag;
            };

            var req = https.request(reqOptions, function (res) {

                var body = '';
                res.setEncoding('utf8');
                res.on('data', function (chunk) { body += chunk; })
                res.on('end', function () {
                    
                    var result = { statusCode: res.statusCode };
                    var type = res.headers['Content-Type'] || res.headers['content-type'];

                    if (type && type.indexOf("xml") > -1) {

                        result.data = domParser.parseFromString(body);
                        return cb(null, result);

                    } else if (type && type.indexOf("json") > -1) {

                        result.data = JSON.parse(body);
                        return cb(null, result);
                    } 
                    else {

                        result.data = body;
                        return cb(null, result);
                    }
                });
            });

            req.end(data);
        }
    };
};

module.exports = Util;
