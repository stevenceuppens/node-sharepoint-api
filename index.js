var Util    = require('./lib/util.js');

var Sharepoint = function (settings) {

    if (!settings || typeof(settings)!=="object") throw new Error("'settings' argument must be a valid object instance.");

    var util = new Util(settings);

    this.authenticate = function(options, cb) {
        util.authenticate(options, cb)
    };

    this.oData = function(options, cb) {
        util.oData(options, cb);
    };

    this.entitySets = function(options, cb) {
        util.entitySets(options, cb);
    };

    this.get = function(options, cb) {
        util.get(options, cb);
    };

    this.query = function(options, cb) {
        util.query(options, cb);
    };

    this.links = function(options, cb) {
        util.links(options, cb);
    };

    this.count = function(options, cb) {
        util.count(options, cb);
    };

    this.create = function(options, cb) {
        util.create(options, cb);
    };

    this.replace = function(options, cb) {
        util.replace(options, cb);
    };

    this.update = function(options, cb) {
        util.update(options, cb);
    };

    this.remove = function(options, cb) {
        util.remove(options, cb);
    };

    util.hook(this);
};

module.exports = Sharepoint;
