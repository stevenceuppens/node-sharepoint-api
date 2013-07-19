var assert  = require("assert");
var API    = require("../index.js");
var mockuire = require("mockuire")(module);

describe("API", function() {

    var settings = {
        host:"sp.com"
    };

    describe("constructor", function () {

        it("should fail on missing setting argument", function (done) {

            try {
                new API();
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings' argument must be a valid object instance.", e.message);
                done();
            }
        });

        it("should fail on invalid setting argument", function (done) {

            try {
                new API("invalid setting");
                throw new Error ("Had to be thrown");
            } catch (e) {
                assert.ok(e);
                assert.ok(e instanceof Error);
                assert.equal("'settings' argument must be a valid object instance.", e.message);
                done();
            }
        });

        it("should work.", function (done) {

            var api = new API({ host: "foo" });
            assert.ok(api);
            assert.ok(api instanceof API);
            done();
        });
    });

    describe("should map all fixed methods", function () {

        var mockedAPI = mockuire("../index", { 
            "./lib/util.js": function () {
                this.hook = function(target) { target.foo = function() {}; };
                this.authenticate = function(options, cb) { cb(null, "authenticate"); };
                this.oData = function(options, cb) { cb(null, "oData"); };
                this.entitySets = function(options, cb) { cb(null, "entitySets"); };
                this.get = function(options, cb) { cb(null, "get"); };
                this.query = function(options, cb) { cb(null, "query"); };
                this.links = function(options, cb) { cb(null, "links"); };
                this.count = function(options, cb) { cb(null, "count"); };
                this.create = function(options, cb) { cb(null, "create"); };
                this.replace = function(options, cb) { cb(null, "replace"); };
                this.update= function(options, cb) { cb(null, "update"); };
                this.remove= function(options, cb) { cb(null, "remove"); };
            }
        });
        
        var api = new mockedAPI({ host: "foo" });

        it("should hook methods", function (done) {
            assert.equal("function", typeof api.foo);
            done();
        });

        it("should map Util's 'oData' method ", function (done) {
            api.oData({}, function(err, val) { 
                assert.ok("oData", val);
                done();
            });
        });

        it("should map Util's 'authenticate' method ", function (done) {
            api.authenticate({}, function(err, val) { 
                assert.ok("authenticate", val);
                done();
            });
        });

        it("should map Util's 'entitySets' method ", function (done) {
            api.entitySets({}, function(err, val) { 
                assert.ok("entitySets", val);
                done();
            });
        });

        it("should map Util's 'get' method ", function (done) {
            api.get({}, function(err, val) { 
                assert.ok("get", val);
                done();
            });
        });

        it("should map Util's 'query' method ", function (done) {
            api.query({}, function(err, val) { 
                assert.ok("query", val);
                done();
            });
        });

        it("should map Util's 'links' method ", function (done) {
            api.links({}, function(err, val) { 
                assert.ok("links", val);
                done();
            });
        });

        it("should map Util's 'count' method ", function (done) {
            api.count({}, function(err, val) { 
                assert.ok("count", val);
                done();
            });
        });

        it("should map Util's 'create' method ", function (done) {
            api.create({}, function(err, val) { 
                assert.ok("create", val);
                done();
            });
        });

        it("should map Util's 'replace' method ", function (done) {
            api.replace({}, function(err, val) { 
                assert.ok("replace", val);
                done();
            });
        });

        it("should map Util's 'update' method ", function (done) {
            api.update({}, function(err, val) { 
                assert.ok("update", val);
                done();
            });
        });

        it("should map Util's 'remove' method ", function (done) {
            api.remove({}, function(err, val) { 
                assert.ok("remove", val);
                done();
            });
        });

    });
});
