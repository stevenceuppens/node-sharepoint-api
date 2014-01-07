var assert  = require("assert");
var Util    = require("../lib/util.js");
var nock    = require("nock");
var fs      = require("fs");

var nowPlusMilliseconds = function(ms) {
    return new Date(new Date().getTime() + ms);
};

describe("Util", function() {

    this.timeout(10000);

    var settings = {
        host:"sp.com"
    };

    var basicAuthSettings = {
        host:"sp.com",
        useBasicAuth: true
    };

    var username = "alfa";
    var password = "beta";

    var successResponseTemplate = "<S:Envelope xmnls:S=\"\" xmnls:wst=\"\" xmnls:wsse=\"\" xmnls:wsu=\"\" xmnls=\"\"><S:Body><wst:RequestSecurityTokenResponse><wst:Lifetime><wsu:Created>{created}</wsu:Created><wsu:Expires>{expires}</wsu:Expires></wst:Lifetime><wst:RequestedSecurityToken><wsse:BinarySecurityToken Id=\"0\">{token}</wsse:BinarySecurityToken></wst:RequestedSecurityToken></wst:RequestSecurityTokenResponse></S:Body></S:Envelope>";
    var successResponse = successResponseTemplate
        .replace("{created}", new Date().toISOString())
        .replace("{expires}", nowPlusMilliseconds(15 * 60000).toISOString()) // in 15 minutes
        .replace("{token}", "authToken");

    var failureResponse = "<S:Envelope xmnls:S=\"\" xmnls:psf=\"\"><S:Body><S:Fault><S:Detail><psf:error><psf:internalerror><psf:text>{error}</psf:text></psf:internalerror></psf:error></S:Detail></S:Fault></S:Body></S:Envelope>"
        .replace("{error}", "failure");

    var samlTemplate = new Util(settings).samlTemplate
        .replace("{endpoint}", "https://sp.com/_forms/default.aspx?wa=wsignin1.0");

    var metadata = fs.readFileSync("./test/metadata.xml");

    beforeEach( function (done) {
        nock.cleanAll();
        done();
    });

    describe("constructor", function () {

        it("should fail on missing setting argument", function (done) {

            try {
                new Util();
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings' argument must be an object instance.", e.message);
                done();
            }
        });

        it("should fail on invalid setting argument", function (done) {

            try {
                new Util("invalid setting");
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings' argument must be an object instance.", e.message);
                done();
            }
        });

        it("should fail when property 'host' is missing.", function (done) {

            try {
                new Util({ });
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings.host' property is a required string.", e.message);
                done();
            }
        });

        it("should fail when property 'host' is invalid.", function (done) {

            try {
                new Util({ host: 1 });
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings.host' property is a required string.", e.message);
                done();
            }
        });

        it("should fail when property 'timeout' is invalid.", function (done) {

            try {
                new Util({ host: "foo", timeout: "invalid" });
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings.timeout' property must be a number.", e.message);
                done();
            }
        });

        it("should fail on invalid username", function (done) {

            try {
                new Util({ host: "foo", username: 10 } );
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings.username' property must be a string.", e.message);
                done();
            }
        });

        it("should fail on invalid password", function (done) {

            try {
                new Util({ host: "foo", password: 10 } );
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings.password' property must be a string.", e.message);
                done();
            }
        });

        it("should work.", function (done) {

            var util = new Util({ host: "foo" });
            assert.ok(util);
            assert.ok(util instanceof Util);
            done();
        });
    });

    describe("authentication",  function() {

        it ("should fail if invalid options", function(done) {

            new Util(settings).authenticate("invalid options", function (err, result) {

                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf("'options'") > -1);

                done(); 
            });
        });

        it ("should fail if username is missing", function(done) {

            new Util(settings).authenticate({ }, function (err, result) {

                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf("'options.username'") > -1);

                done(); 
            });
        });

        it ("should fail if password is missing", function(done) {

            new Util(settings).authenticate({ username: username }, function (err, result) {

                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf("'options.password'") > -1);

                done(); 
            });
        });

        it ("should fail if password is invalid", function(done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", "invalid password");

            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, failureResponse);

            new Util(settings).authenticate({ username: username, password: "invalid password" }, function (err, result) {

                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("failure", err.message);

                loginNock.done();
                done(); 
            });
        });

        it ("should authenticate", function(done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse);

            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            new Util(settings).authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);
                assert.equal('string', typeof result.auth);
                assert.equal(36, result.auth.length);

                loginNock.done();
                authzNock.done();
                metadataNock.done();
                done(); 
            });
        });

        it ("should cache authentication", function(done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse);

            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            var util = new Util(settings);
            util.authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);
                assert.equal('string', typeof result.auth);
                assert.equal(36, result.auth.length);
    
                util.authenticate({ username: username, password: password }, function (err2, result2) {

                    assert.ok(!err2);
                    assert.ok(result2);
                    assert.equal(result.auth, result2.auth);

                    loginNock.done();
                    authzNock.done();
                    metadataNock.done();
                    done(); 
                });
            });
        });

        it ("should drop expired items from auth cache", function(done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            // will authenticate twice because item got expired
            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse)
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse);

            // will authorize twice because item got expired
            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] })
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

            // Metadata will be retrieved just once.
            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            var timeout = 100;

            var util = new Util({ host: settings. host, timeout: timeout });
            util.authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);
                assert.equal('string', typeof result.auth);
                assert.equal(36, result.auth.length);
    
                setTimeout( function() {
                    util.authenticate({ username: username, password: password }, function (err2, result2) {

                        assert.ok(!err2);
                        assert.ok(result2);
                        assert.equal('string', typeof result2.auth);
                        assert.equal(36, result2.auth.length);
                        assert.ok(result.auth !== result2.auth);

                        loginNock.done();
                        authzNock.done();
                        metadataNock.done();
                        done(); 
                    });
                }, timeout + 100);
            });
        });

        it ("should renew expired items tokens", function(done) {
            this.timeout(5000);

            var expirationMilliseconds = 2000;
            var firstTokenExpiresOn = nowPlusMilliseconds(expirationMilliseconds + 60000); // I'm adding a minute because the util.js removes a minute for avoid time conflict
            
            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            // will authorize twice because token got expired
            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authTokenExpiresSoon")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] })  // first authz
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=rst", "rtFa=uvw"] }); // second authz

            
            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')  // Should match first authz
                .get("/_api/$metadata")                         // Metadata will be retrieved just once.
                .reply(200, metadata, { "content-type": "application/xml" })
                .get("/_api/Lists?$inlinecount=none")           // invocation to query list
                .reply(200, {foo: "bar"}, { "content-type": "application/json" });

            var spNock2 = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=rst;rtFa=uvw')  // Should match second authz
                .get("/_api/Users?$inlinecount=none")           // invocation to query users
                .reply(200, {foo: "baz"}, { "content-type": "application/json" });

            // will authenticate twice because item got expired
            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, function() {
                    // returns a token that expires en 200 ms
                    return successResponseTemplate
                        .replace("{created}", new Date().toISOString())
                        .replace("{expires}", firstTokenExpiresOn.toISOString())
                        .replace("{token}", "authTokenExpiresSoon");
                })
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse); // returns a normal token

            var util = new Util(settings);
            util.authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);
                assert.equal('string', typeof result.auth);
                assert.equal(36, result.auth.length);
    
                util.query({ auth: result.auth, resource: "Lists" }, function(err, result2) {

                    assert.ok(!err);
                    assert.ok(result2);
                    assert.equal("bar", result2.data.foo);

                    setTimeout( function() {

                        util.query({ auth: result.auth, resource: "Users" }, function(err, result3) {

                            assert.ok(!err);
                            assert.ok(result3);
                            assert.equal("baz", result3.data.foo);

                            loginNock.done();
                            authzNock.done();
                            spNock.done();
                            spNock2.done();
                            done(); 
                        });
                    }, expirationMilliseconds);
                });
            });
        });

        it ("should cache the entity sets", function(done) {

            var saml = samlTemplate
                .replace("{username}", "foo")
                .replace("{password}", password);

            var saml2 = samlTemplate
                .replace("{username}", "bar")
                .replace("{password}", password);

            var successResponse2 = successResponseTemplate.replace("{token}", "authToken2");

            // should have two authentications
            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml) 
                .reply(200, successResponse)
                .post("/extSTS.srf", saml2)  
                .reply(200, successResponse2);

            // should have two authorizations
            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] })
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken2")
                .reply(200, "", { "set-cookie": ["FedAuth=alfa", "rtFa=beta"] });

            // should have only one request for metadata
            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            var util = new Util(settings);

            util.authenticate({ username: "foo", password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);

                util.authenticate({ username: "bar", password: password }, function (err, result) {

                    assert.ok(!err);
                    assert.ok(result);

                    loginNock.done();
                    authzNock.done();
                    metadataNock.done();
                    done(); 
                });
            });
        });
    });

    describe("basic authentication",  function() {

        it ("should fail if password is invalid", function(done) {

            var loginNock = new nock("https://sp.com")
            .get("/_api/lists")  
            .reply(401, failureResponse);

            new Util(basicAuthSettings).authenticate({ username: username, password: "invalid password" }, function (err, result) {

                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("Authentication fail.", err.message);

                loginNock.done();
                done(); 
            });
        });

        it ("should authenticate using basic auth", function(done) {

            var loginNock = new nock("https://sp.com")
            .matchHeader('Authorization', 'Basic ' + new Buffer(username + ':' + password).toString('base64'))
            .get("/_api/lists")  
            .reply(200, successResponse);

            new Util(basicAuthSettings).authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);
                assert.equal('string', typeof result.auth);
                assert.equal(36, result.auth.length);

                loginNock.done();
                done(); 
            });
        });
    });

    describe("oData method", function() {

        var util;
        var username = "alfa";
        var password = "beta";

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if not auth value or user's credentials were passed within options argument.", function ( done ) {

            util.oData({ command: "foo" }, function(err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf('options.username') > -1);
                done();
            });
        });

        it ("should fail if 'command' property is missing.", function ( done ) {

            util.oData({ auth: "xyz" }, function(err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf('options.command') > -1);
                done();
            });
        });

        it ("should fail if 'command' property is invalid.", function ( done ) {

            util.oData({ auth: "xyz", command: 1 }, function(err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.ok(err.message.indexOf('options.command') > -1);
                done();
            });
        });

        it ("should be able to invoke a method passing user's credentials" , function (done) {

            util.authenticate = function(options, cb) {
                
                util.cacheAuth.set("xyz", {
                    authz       : { FedAuth:"alfa", rtFa:"beta"},
                    cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                    username    : username,
                    password    : password
                });

                cb(null, { auth: "xyz"} );
            };

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=alfa;rtFa=beta')
                .get("/_api/foo")
                .reply(200, {foo:"bar"}, { "content-type": "application/json" });

            util.oData({ username: username, password: password, command: "foo" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result.data);
                assert.equal("bar", result.data.foo);

                spNock.done();
                done();
            });
        });

        it ("should be able to invoke a method passing auth value" , function (done) {

            var item = {
                authz       : { FedAuth:"alfa", rtFa:"beta"},
                cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                username    : username,
                password    : password
            };

            util.cacheAuth.set("xyz", item);

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', item.cookieAuthz)
                .get("/_api/foo")
                .reply(200, {foo:"bar"},{ "content-type": "application/json" });

            util.oData({ auth: "xyz", command: "foo" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result.data);
                assert.equal("bar", result.data.foo);
                done();
            });
        });

        it ("should be able to invoke a method using a different HTTP method" , function (done) {

            var item = {
                authz       : { FedAuth:"alfa", rtFa:"beta"},
                cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                username    : username,
                password    : password
            };

            util.cacheAuth.set("xyz", item);

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', item.cookieAuthz)
                .post("/_api/foo")
                .reply(200, {foo:"bar"},{ "content-type": "application/json" });

            util.oData({ auth: "xyz", command: "foo",method: "POST" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result.data);
                assert.equal("bar", result.data.foo);
                done();
            });
        });

        it ("should send 'data' argument as request's body" , function (done) {

            var item = {
                authz       : { FedAuth:"alfa", rtFa:"beta"},
                cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                username    : username,
                password    : password
            };

            util.cacheAuth.set("xyz", item);

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', item.cookieAuthz)
                .post("/_api/foo", {baz:1})
                .reply(200, {foo:"bar"},{ "content-type": "application/json" });

            util.oData({ auth: "xyz", command: "foo", method: "POST", data: { baz: 1 } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result.data);
                assert.equal("bar", result.data.foo);
                done();
            });
        });

        it ("should send 'etag' argument as request's header" , function (done) {

            var item = {
                authz       : { FedAuth:"alfa", rtFa:"beta"},
                cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                username    : username,
                password    : password
            };

            util.cacheAuth.set("xyz", item);

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', item.cookieAuthz)
                .matchHeader('if-match', "W")
                .get("/_api/foo")
                .reply(200, {foo:"bar"},{ "content-type": "application/json" });

            util.oData({ auth: "xyz", command: "foo", etag: "W" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result.data);
                assert.equal("bar", result.data.foo);
                done();
            });
        });
    });

    describe("entitySets method", function() {

        var util;
        var username = "alfa";
        var password = "beta";

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });


        it("should get the list of entity sets passing as options the auth token.", function (done) {

            util.cacheAuth.set("xyz", {
                authz       : { FedAuth:"alfa", rtFa:"beta"},
                cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                username    : username,
                password    : password
            });

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=alfa;rtFa=beta')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            util.entitySets({ auth: "xyz" }, function (err, result) {
                
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result instanceof Array);
                assert.ok(result.indexOf("Lists") > -1);

                spNock.done();
                done();
            });
        });


        it("should get the list of entity sets passing as options the user credentials.", function (done) {

            util.authenticate = function(options, cb) {
                
                util.cacheAuth.set("xyz", {
                    authz       : { FedAuth:"alfa", rtFa:"beta"},
                    cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                    username    : username,
                    password    : password
                });

                cb(null, { auth: "xyz"} );
            };

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=alfa;rtFa=beta')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });

            util.entitySets({ username: username, password: password }, function (err, result) {
                
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result instanceof Array);
                assert.ok(result.indexOf("Lists") > -1);

                spNock.done();
                done();
            });
        });



        it("should return an empty array if '$metadata'is not supported by sharepoint server.", function (done) {

            util.authenticate = function(options, cb) {
                
                util.cacheAuth.set("xyz", {
                    authz       : { FedAuth:"alfa", rtFa:"beta"},
                    cookieAuthz : 'FedAuth=alfa;rtFa=beta',
                    username    : username,
                    password    : password
                });

                cb(null, { auth: "xyz"} );
            };

            var spNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=alfa;rtFa=beta')
                .get("/_api/$metadata")
                .reply(404);

            util.entitySets({ username: username, password: password }, function (err, result) {
                
                assert.ok(!err);
                assert.ok(result);
                assert.ok(result instanceof Array);
                assert.equal(result.length, 0);

                spNock.done();
                done();
            });
        });
    });

    describe("get method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.get({ id: "bar" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.get({ resource: 10, id: "bar" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'id' property is missing", function (done) {
            util.get({ resource: "foo" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'id' property is missing.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo('bar')", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.get({resource: "foo", id:"bar"}, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo(123)", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.get({resource: "foo", id: 123}, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("query method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.query({ }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.get({ resource: 10 }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should invoke oData", function (done) {

            util.oData = function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo?$filter=bar&$expand=baz&$select=xyz&$orderby=pqr&$top=rst&$skip=uvw&$inlinecount=allpages", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            var options = {
                resource    : "foo", 
                filter      : "bar", 
                expand      : "baz", 
                select      : "xyz", 
                orderBy     : "pqr", 
                top         : "rst", 
                skip        : "uvw", 
                inLineCount : true
            };

            util.query(options, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with default values", function (done) {

            util.oData = function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo?$inlinecount=none", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.query({ resource: "foo" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("links method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.links({ id: "bar", entity: "baz" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.links({ resource: 10, id: "bar", entity: "baz" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'id' property is missing", function (done) {
            util.links({ resource: "foo", entity: "baz" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'id' property is missing.", err.message);
                done();
            });
        });

        it ("should fail if 'entity' property is missing", function (done) {
            util.links({ resource: "foo", id: "bar" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'entity' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'entity' property is invalid", function (done) {
            util.links({ resource: "foo", id: "bar", entity: 10 }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'entity' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo('bar')/$links/baz", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.links({resource: "foo", id:"bar", entity: "baz" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo(123)/$links/baz", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.links({resource: "foo", id:123, entity: "baz" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("count method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.count({ id: "bar", entity: "baz" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.count({ resource: 10, id: "bar", entity: "baz" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo/$count", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.count({resource: "foo" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("GET", options.method);
                assert.equal ("/foo(123)/$links/baz", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.links({resource: "foo", id:123, entity: "baz" }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("create method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.create({ data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.create({ resource: 10, data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should invoke oData", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("POST", options.method);
                assert.equal ("/foo", options.command);
                assert.ok    (options.data);
                assert.equal ("baz", options.data.bar);
                cb(null, {statusCode: 200, data: true});
            };

            util.create({ resource: "foo", data: { bar: "baz" } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("replace method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.replace({ id: 10, data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.replace({ id: 10, resource: 10, data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'id' property is missing", function (done) {
            util.replace({ resource: "foo", data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'id' property is missing.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("PUT", options.method);
                assert.equal ("/foo('xyz')", options.command);
                assert.ok    (options.data);
                assert.equal ("baz", options.data.bar);
                cb(null, {statusCode: 200, data: true});
            };

            util.replace({ resource: "foo", id: "xyz", data: { bar: "baz" } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("PUT", options.method);
                assert.equal ("/foo(123)", options.command);
                assert.ok    (options.data);
                assert.equal ("baz", options.data.bar);
                cb(null, {statusCode: 200, data: true});
            };

            util.replace({ resource: "foo", id: 123, data: { bar: "baz" } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("update method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.update({ id: 10, data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.update({ id: 10, resource: 10, data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'id' property is missing", function (done) {
            util.update({ resource: "foo", data: { bar: "baz" } }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'id' property is missing.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("PATCH", options.method);
                assert.equal ("/foo('xyz')", options.command);
                assert.ok    (options.data);
                assert.equal ("baz", options.data.bar);
                cb(null, {statusCode: 200, data: true});
            };

            util.update({ resource: "foo", id: "xyz", data: { bar: "baz" } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("PATCH", options.method);
                assert.equal ("/foo(123)", options.command);
                assert.ok    (options.data);
                assert.equal ("baz", options.data.bar);
                cb(null, {statusCode: 200, data: true});
            };

            util.update({ resource: "foo", id: 123, data: { bar: "baz" } }, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });

    describe("remove method", function() {

        var util;

        beforeEach(function (done) {
            util = new Util(settings);
            done();
        });

        it ("should fail if 'resource' property is missing", function (done) {
            util.remove({ id: "bar" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'resource' property is invalid", function (done) {
            util.remove({ resource: 10, id: "bar" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'resource' property is missing or invalid.", err.message);
                done();
            });
        });

        it ("should fail if 'id' property is missing", function (done) {
            util.remove({ resource: "foo" }, function (err, result) {
                assert.ok(err);
                assert.ok(err instanceof Error);
                assert.equal("'id' property is missing.", err.message);
                done();
            });
        });

        it ("should invoke oData with a string as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("DELETE", options.method);
                assert.equal ("/foo('bar')", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.remove({resource: "foo", id:"bar"}, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });

        it ("should invoke oData with a number as id", function (done) {
            util.oData= function (options, cb) {
                assert.equal ("DELETE", options.method);
                assert.equal ("/foo(123)", options.command);
                cb(null, {statusCode: 200, data: true});
            };

            util.remove({resource: "foo", id: 123}, function (err, result) {
                assert.ok(!err);
                assert.ok(result);
                assert.equal(200, result.statusCode);
                assert.equal(true, result.data);
                done();
            });
        });
    });


    describe("hook method", function() {

        it ("should add methods after an user was authenticated", function (done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse);

            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });


            var target = {};
            var util = new Util(settings);

            util.hook(target);

            assert.ok(!target.getLists);

            util.authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);

                assert.equal("function", typeof target.getLists);

                loginNock.done();
                authzNock.done();
                metadataNock.done();
               done(); 
            });
        });


        it ("should add methods if an user already was authenticated", function (done) {

            var saml = samlTemplate
                .replace("{username}", username)
                .replace("{password}", password);

            var loginNock = new nock("https://login.microsoftonline.com")
                .post("/extSTS.srf", saml)  
                .reply(200, successResponse);

            var authzNock = new nock("https://sp.com")
                .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

            var metadataNock = new nock("https://sp.com")
                .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                .get("/_api/$metadata")
                .reply(200, metadata, { "content-type": "application/xml" });


            var util = new Util(settings);
            util.authenticate({ username: username, password: password }, function (err, result) {

                assert.ok(!err);
                assert.ok(result);

                var target = {};
                util.hook(target);
                assert.equal("function", typeof target.getLists);

                loginNock.done();
                authzNock.done();
                metadataNock.done();
                done(); 
            });
        });

        it ("Should map correct methods", function() {

            var target = {};

            before (function (done){
                var saml = samlTemplate
                    .replace("{username}", username)
                    .replace("{password}", password);

                var loginNock = new nock("https://login.microsoftonline.com")
                    .post("/extSTS.srf", saml)  
                    .reply(200, successResponse);

                var authzNock = new nock("https://sp.com")
                    .post("/_forms/default.aspx?wa=wsignin1.0", "authToken")
                    .reply(200, "", { "set-cookie": ["FedAuth=xyz", "rtFa=pqr"] });

                var metadataNock = new nock("https://sp.com")
                    .matchHeader('cookie', 'FedAuth=xyz;rtFa=pqr')
                    .get("/_api/$metadata")
                    .reply(200, metadata, { "content-type": "application/xml" });


                var util = new Util(settings);
                util.get = function(options, cb) { cb(null, "get-" + options.resource); };
                util.query = function(options, cb) { cb(null, "query-" + options.resource); };
                util.links = function(options, cb) { cb(null, "links-" + options.resource); };
                util.count = function(options, cb) { cb(null, "count-" + options.resource); };
                util.create = function(options, cb) { cb(null, "create-" + options.resource); };
                util.replace = function(options, cb) { cb(null, "replace-" + options.resource); };
                util.update = function(options, cb) { cb(null, "update-" + options.resource); };
                util.remove = function(options, cb) { cb(null, "remove-" + options.resource); };

                util.authenticate({ username: username, password: password }, function (err, result) {

                    assert.ok(!err);
                    assert.ok(result);

                    util.hook(target);
                    
                    target.getLists({}, function(err, res){
                        assert.ok(!err);
                        assert.equal("get-Lists", res);

                        target.queryLists({}, function(err, res){
                            assert.ok(!err);
                            assert.equal("query-Lists", res);

                            target.linksLists({}, function(err, res){
                                assert.ok(!err);
                                assert.equal("links-Lists", res);

                                target.countLists({}, function(err, res){
                                    assert.ok(!err);
                                    assert.equal("count-Lists", res);

                                    target.createLists({}, function(err, res){
                                        assert.ok(!err);
                                        assert.equal("create-Lists", res);

                                        target.replaceLists({}, function(err, res){
                                            assert.ok(!err);
                                            assert.equal("replace-Lists", res);

                                            target.updateLists({}, function(err, res){
                                                assert.ok(!err);
                                                assert.equal("update-Lists", res);

                                                target.removeLists({}, function(err, res){
                                                    assert.ok(!err);
                                                    assert.equal("remove-Lists", res);

                                                    loginNock.done();
                                                    authzNock.done();
                                                    metadataNock.done();
                                                    done(); 
                                                });
                                            });
                                        });
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    });
});

