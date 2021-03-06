// module dependencies
require("simple-errors");
var https       = require('https');
var url         = require('url');
var xpath       = require('xpath');
var cookie      = require('cookie');
var Cache       = require("mem-cache");
var uuid        = require("node-uuid");
var domParser   = new (require('xmldom').DOMParser)();
var constants   = require('constants');

// this class implements all features 
var Util = function (settings) {

    // Arguments validation
    if (!settings || typeof(settings)!=="object") throw new Error("'settings' argument must be an object instance.");
    if (!settings.host || typeof(settings.host)!=="string") throw new Error("'settings.host' property is a required string.");
    if (settings.timeout!==undefined && typeof(settings.timeout)!=="number") throw new Error("'settings.timeout' property must be a number.");
    if (settings.username && typeof(settings.username)!=="string") throw new Error("'settings.username' property must be a string.");
    if (settings.password && typeof(settings.password)!=="string") throw new Error("'settings.password' property must be a string.");

    // use basic auth
    settings.useBasicAuth = settings.useBasicAuth || false;

    var self            = this;     // Auto reference
    var entitySets      = null;     // String array containing all entity sets names
    var pendingHook     = null;     // Function that will be executed after 'entitySets' array was populated.
    var loginPath       = '/_forms/default.aspx?wa=wsignin1.0';                 // Login path 
    var loginEndpoint   = url.resolve("https://" + settings.host, loginPath);   // Login URL for the configured host

    // Sets default arguments values
    settings.timeout = settings.timeout || 15 * 60 * 1000;  // default sessions timeout of 15 minutes in ms   

    // Cache by authentication token, containing all session instances
    Object.defineProperty(this, "cacheAuth", {
        enumerable: false,
        configurable: false,
        writable: false,
        value: new Cache(settings.timeout)
    });

    // cache by user name, containing all authentication tokens
    var cacheUser   = new Cache(settings.timeout);   // Cache by auth tokens 

    // templete of the SAML token 
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

    // authenticates and authorizes the user, and stores the session into the cache
    this.authenticate = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        var authOptions = clone(options) || {};

        // validates arguments values
        if (typeof authOptions !== 'object') return cb(new Error("'options' argument is missing or invalid."));

        // Validates username and password 
        authOptions.username = authOptions.username || settings.username;
        authOptions.password = authOptions.password || settings.password;
        authOptions.useBasicAuth = authOptions.useBasicAuth || settings.useBasicAuth;

        if (!(authOptions.username)) return cb(new Error("'options.username' property is required."));
        if (!(authOptions.password)) return cb(new Error("'options.password' property is required."));

        // tries to find an existing session for the user
        var auth = cacheUser.get(authOptions.username);

        if (auth) {
            // returns session info if both passwords do match
            var item = self.cacheAuth.get(auth);
            if (item && item.password === authOptions.password && new Date() < item.expires) {

                return cb(null, { 
                    auth: auth, 
                    user: item.username, 
                    authz: item.authz,
                    authHeader: item.authHeader 
                });
            }

            self.cacheAuth.remove(auth);
            cacheUser.remove(authOptions.username);
        }

        if (authOptions.useBasicAuth) {
            // creates a new authId
            auth = uuid.v4();

            // creates cache item (session)
            var item = {
                authHeader : 'Basic ' + new Buffer(authOptions.username + ':' + authOptions.password).toString('base64'),
                username    : authOptions.username,
                password    : authOptions.password
            };

            // stores session 
            self.cacheAuth.set(auth, item);
            cacheUser.set(authOptions.username, auth);

            // Verifies that credentials are valid and have access rights.
            verifyAccess(auth, function(err) {
                if (err) {
                    // cleans caches
                    if (auth) {
                        self.cacheAuth.remove(auth);
                        cacheUser.remove(authOptions.username)
                    }

                    return cb(err);
                }

                cb(null, { auth: auth, user: authOptions.username });
            });

            return;
        };

        // no session was found. 
        // Authenticates the user
        getAuthenticationToken(authOptions.username, authOptions.password, function (err, authToken) {

            // validates error
            if (err) {
                // cleans caches
                if (auth) {
                    self.cacheAuth.remove(auth);
                    cacheUser.remove(authOptions.username)
                }
                return cb(err);
            }

            // Authorize the user
            getAuthorization(authToken.token, function (err, authz) {

                // Validates error
                if (err) {
                    // cleans caches
                    if (auth) {
                        self.cacheAuth.remove(auth);
                        cacheUser.remove(authOptions.username)
                    }
                    return cb(err);
                }

                // reuse auth by username
                auth = auth || uuid.v4();

                // creates cache item (session)
                var item = {
                    expires     : authToken.expires,
                    authz       : authz,
                    cookieAuthz : 'FedAuth=' + authz.FedAuth + ';rtFa=' + authz.rtFa,
                    username    : authOptions.username,
                    password    : authOptions.password
                };

                // stores session 
                self.cacheAuth.set(auth, item);
                cacheUser.set(authOptions.username, auth);
                
                // populates entitySets array
                getEntitySets(auth, function(err) {

                    // validates error
                    if (err) return cb(err);
                    cb(null, { auth: auth, user: authOptions.username });
                });
            });
        });
    };


    // Executes a oData Command
    this.oData = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options || typeof options !== 'object')                    return cb(new Error("'options' argument is missing or invalid."));
        if (!options.command || typeof options.command !== 'string')    return cb(new Error("'options.command' argument is missing or invalid."));

        request(options, cb);
    };


    // Returns an string array with the name of all entity sets
    this.entitySets = function(options, cb) {
        
        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options || typeof options !== 'object') return cb(new Error("'options' argument is missing or invalid."));

        if (options.auth) {

            getEntitySets(options.auth, cb);
        } else {

            this.authenticate(options, function(err, result) {
                if (err) return cb(new Error("Couldn't get the entity sets. " + err));
                getEntitySets(result.auth, cb);
            });
        }
    };


    // gets an entity by its id
    this.get = function(options, cb) {
        
        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // executes a query on an entity set
    this.query = function(options, cb) {
        
        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // gets all links of an entity instance to entities of a specific type 
    this.links = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // returns the number of elements of an entity set
    this.count = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // adds an entity instance to an entity set
    this.create = function(options, cb) {
        
        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // does a partial update of an existing entity instance 
    this.replace = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // does a complete update of an existing entity instance
    this.update = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // removes an existing instance from an entity set
    this.remove = function(options, cb) {

        // handles optional 'options' argument
        if (!cb && typeof options === 'function') {
            cb = options;
            options = {};
        }

        // sets default values
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


    // adds methods dinamically to the instance passed by parameter.
    // for each entity set, methods for get, update, create, etc. will be added.
    this.hook = function(target) {

        // function for add a single method
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

        // function for add all methods to a every entity set
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

        // adds te methods if the array of entity sets was populated
        if (entitySets) {
            addAllMethods(entitySets);
        } else {
            // wait for the entitySets
            pendingHook = function() {
                addAllMethods(entitySets);
                pendingHook = null;
            };
        }
    };

    // returns 
    this.lookupMethod = function(connectorInstance, name, cb) {

        if (!pendingHook) {
            var method = connectorInstance[name];
            if (typeof method === 'function') return cb(null, method);
            return cb();
        }

        // returns a wrappers that waits until entity sets were retrieved
        var wrapper = function (options, callback) {

            // forces the invocation to entity sets
            self.entitySets(options, function (err, result) {
                if (err) return callback(err);

                // after entitySets were retrieved, invokes the method or returns not found (404)
                var method = connectorInstance[name];
                if (typeof method === 'function') return method(options, callback);
                return callback(null, null);

            });
        };

        return cb(null, wrapper);
    };

    // verifies access
    var verifyAccess = function (auth, cb) {
        
        var options = {
            command: 'lists',
            auth: auth
        }

        request(options, function(err, results) {
            if (err || results.statusCode == 401 || results.statusCode == 403) {
                cb(err || new Error('Authentication fail.'));
                return;
            };

            cb();
        })
    }

    // serialize the entity id
    var buildODataId = function(id) {

        if (typeof id === 'string') return "'" + id + "'";
        return "" + id;        
    }

    // retrives the entity sets from the service's metadata
    var getEntitySets = function(auth, cb) {

        // retrieves entity sets only once
        if (entitySets) return cb(null, entitySets);

        // gets the metadata XML
        self.oData({ auth: auth, command: "$metadata"}, function(err, result) {

            try {
                if (err) {
                    // If server does not support $metadata, the entityset will be setted to an empty string 
                    entitySets = [];
                } else {
                    // gets the entity sets names using a XPath query
                    entitySets = xpath
                        .select("//*[local-name(.)='EntitySet' and namespace-uri(.)='http://schemas.microsoft.com/ado/2009/11/edm']/@Name", result.data)
                        .map(function(attr) { return attr.value;} );
                }
                // adds methods dinamically
                if (pendingHook) pendingHook();

                // returns entities
                cb(null, entitySets);

            } catch(e) {

                // process error
                entitySets = null;
                cb(new Error("Couldn't get entity sets. " + e));
            }
        });  
    }

    // build full name condition for XPath expression
    var name = function(name) {
        return "/*[name(.)='" + name + "']";
    };

    // does the HTTP POST and returns the authentication token
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
                
                // parses XML
                var resXml = domParser.parseFromString(xml); 
                
                // search for a fault
                var exp = ['S:Envelope', 'S:Body', 'S:Fault', 'S:Detail', 'psf:error', 'psf:internalerror', 'psf:text'].map(name).join("") + "/text()";
                var fault = xpath.select(exp, resXml);
                if (fault.length > 0) return cb(new Error(fault.toString()));
        
                // get expiration time
                var expires = null;
                var expiresText = null;
                exp = ['S:Envelope', 'S:Body', 'wst:RequestSecurityTokenResponse', 'wst:Lifetime', 'wsu:Expires'].map(name).join("") + "/text()";

                try {
                    expiresText = xpath.select(exp, resXml).toString();
                    expires = new Date(new Date(expiresText).getTime() - 60 * 1000); // substract a minute from expiration time, to avoid limit conflict.
                } catch (e) {
                    console.log ("WARNING: Couldn't parse token's expiration time. Error: " + e + " Text value: " + expiresText);
                    console.log ("\tToken's expiration time was set to 15 minutes.");
                    exires = new Date(new Date().getTime() + 15 * 60 * 1000);      // by default expires in 15 minutes. 
                }

                // gets auth token
                exp = ['S:Envelope', 'S:Body', 'wst:RequestSecurityTokenResponse', 'wst:RequestedSecurityToken', 'wsse:BinarySecurityToken'].map(name).join("") + "/text()";
                var token = xpath.select(exp, resXml);
                if (token.length > 0) return cb(null, { token: token.toString(), expires: expires });

                cb(new Error("Invalid empty token was received"));
            })
        });
        
        req.end(samlRequest);
    };

    // does the HTTP POST and returns the authorization data
    var getAuthorization = function(authToken, cb) {

        var options = {
            method: 'POST',
            host: settings.host,
            path: loginPath,
            secureOptions: constants.SSL_OP_NO_TLSv1_2,
            ciphers: 'ECDHE-RSA-AES256-SHA:AES256-SHA:RC4-SHA:RC4:HIGH:!MD5:!aNULL:!EDH:!AESGCM',
            honorCipherOrder: true,
            agent: false,
            headers: {
                'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0)',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Content-Length': Buffer.byteLength(authToken)
            }
        };

        var req = https.request(options, function (res) {

            res.setEncoding('utf8');
            res.once('data', function () {
                
                var cookies = cookie.parse(res.headers["set-cookie"].join(";"));
                cb (null, {
                    FedAuth : cookies.FedAuth,
                    rtFa    : cookies.rtFa                    
                });
            });
        });

        req.write(authToken);
        req.end();
    };


    var buildPath = function(command) {
        var path = (settings.site ? "/" + settings.site : "") + "/_api/" + command;

        while (path.indexOf("//") >- 1) {
            path = path.replace("//", "/");
        }

        return encodeURI(path);
    };

    // sends oData requests to the server
    var request = function(options, cb) {

        if (!options.auth) {
            // tries to authenticate the user  because the auth token was passed.
            self.authenticate({ username: options.username, password: options.password }, function (err, result) {
                if (err) return cb(err);

                // tries again
                options.auth = result.auth;
                request(options, cb);
            });
            return;
        } 

        // gets cached session data
        var item = self.cacheAuth.get(options.auth);
        if (!item) return cb (new Error("Invalid 'auth' property."));

        if (!settings.useBasicAuth) {
            // is the session expired?
            if ( new Date() >= item.expires) {

                // token is expired, gets a new one
                self.authenticate( {username: item.username, password: item.password}, function (err, result) {
                    if (err) return cb(err);

                    // tries again
                    options.auth = result.auth;
                    request(options, cb);
                });
                return;
            }

        };

        //Builds request options
        var data = options.data ? JSON.stringify(options.data) : "";

        var reqOptions = {
            method  : options.method || "GET",
            host    : settings.host,
            path    : buildPath(options.command),
            secureOptions: constants.SSL_OP_NO_TLSv1_2,
            ciphers: 'ECDHE-RSA-AES256-SHA:AES256-SHA:RC4-SHA:RC4:HIGH:!MD5:!aNULL:!EDH:!AESGCM',
            honorCipherOrder: true,
            agent: false,
            headers : {
                'accept'        : 'application/json;odata=verbose',
                'content-type'  : 'application/json',
                'content-length': data.length
            }
        };

        if (settings.useBasicAuth) {
            reqOptions.headers['Authorization'] = item.authHeader;
        }
        else {
            reqOptions.headers['cookie'] = item.cookieAuthz;
        }

        if (options.etag) {
            reqOptions.headers['if-match'] = options.etag;
        };

        // Builds HTTP request
        var req = https.request(reqOptions, function (res) {

            var body = '';
            res.setEncoding('utf8');
            res.on('data', function (chunk) { body += chunk; })
            res.on('end', function () {
                var result = { statusCode: res.statusCode };
                var type = res.headers['Content-Type'] || res.headers['content-type'];

                if (type && type.indexOf("xml") > -1) {

                    // returns parsed XML
                    result.data = domParser.parseFromString(body);
                    return cb(null, result);

                } else if (type && type.indexOf("json") > -1) {

                    // returns parsed JSON
                    result.data = JSON.parse(body);
                    return cb(null, result);
                } 
                else {

                    // returns plain body 
                    result.data = body;
                    return cb(null, result);
                }
            });
        });

        // Sends HTTP request
        req.end(data);
    };

    // Reeturns a new copy of an object or value
    var clone =  function(source) {
    
        // is source null or a value type?
        if (source === null || typeof source !== 'object') return source;

        // returns a copy of an array
        if (source instanceof Array)    return source.map(clone);

        // returns a copy of a date
        if (source instanceof Date)     return new Date(source.getTime());

        // returns a copy of an object
        var result = {};
        Object.keys(source).map(function(prop) { result[prop] = clone(source[prop]); });
        return result;
    };
};

module.exports = Util;
