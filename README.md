# SharePoint client for Nodejs
This node module provides a set of methods to interact with SharePont Online services.
The module supports Claims-based authentication and contains methods to query and manage entities like lists, documents, etc.
The module was created as part of [KidoZen](http://www.kidozen.com) project, as a connector for its Enterprise API feature.
## Installation

Use npm to install the module:

```
> npm install sharepoint-api
```

## API

Due to the asynchrounous nature of Nodejs, this module uses callbacks in requests. All callbacks have 2 arguments: `err` and `data`.

```
function callback (err, data) {
	// err contains an Error class instance, if any
	// data contains the resulting data
} 
``` 

### Constructor

The module exports a class and its constructor requires a configuration object with following properties

* `host`: Required string. SharePoint's host name. By instance: foo.sharepoint.com 
* `timeout`: Optional integer for the session timeout in milleseconds. Default 15 minutes.  
* `username`: Optional SharePoint Online's user name.
* `password`: Optional user's password.

```
var SharePoint = require("sharepoint-api");
var sharepoint = new SharePoint({ 
	host: "myOnlineInstance.sharepoint.com", 
	timeout: 5*60*1000 	// Timeout of 5 minutes
});
```

### Methods
All public methods has the same signature. The signature has two arguments: `options` and `callback`.
* `options` must be an object instance containig all parameters for the method.
* `callback` must be a function.

#### authenticate(options, callback)

This method should be used to authenticate user's credentials. A successed authentication will return an object instance containing the `auth` property. The value of this property is the authentication token that will be required by other methods.

**Parameters:**
* `options`: A required object instance containing authentication's parameters:
	* `username`: String.
	* `password`: String.
* `callback`: A required function for callback.

```
sharepoint.authenticate({ username:"foo", password: "bar" }, function(err, result) {
	if (err) return console.error (err);
	console.log (result.auth);
});
```

#### oData(options, callback)

Executes a oData command.
**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `command`: String. Required. oData expression. By instance "/entitySet('id')"
	* `method`: String. Optional. HTTP request method to send. Default value is "GET"
	* `data`: Any type. Optional. HTTP Request's body 
	* `etag`: String. Optional. ETag value for concurrency control.
* `callback`: A required function for callback.

```
// Gets all list instances.
sharepoint.oData({ auth: "...", command: "/Lists" }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### entitySets(options, callback)

Retrieves the list of entity sets.

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
* `callback`: A required function for callback.

```
// Gets all list instances.
sharepoint.entitySets({ auth: "..." }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### get(options, callback)

Gets an entity instance by its id.

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set.
	* `id`: String or number. Required. Id of the entity.
* `callback`: A required function for callback.

```
// Gets the list with id equal to one.
sharepoint.get({ auth: "...", resource:"Lists", id:1 }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### query(options, callback)

Returns matched entity instances from an entity set.

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set.
    * `filter`: String. Optional. oData $filter expression.
    * `expand`: String. Optional. oData $expand expression.
    * `select`: String. Optional. oData $select expression.
    * `orderBy`: String. Optional. oData $orderby expression.
    * `top`: String. Optional. oData $top expression.
    * `skip`: String. Optional. oData $skip expression.
    * `inLineCount`: Boolean. Optional. Default value is false. 
* `callback`: A required function for callback.

```
// Query for all private lists.
sharepoint.query({ auth: "...", resource:"Lists", filter:"IsPrivate eq true" }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### links(options, callback)

Retrieves links between entities

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set of the source entity.
	* `id`: String or number. Required. Id of the source entity.
	* `entity`: String or number. Required. Name of the target entity type.
* `callback`: A required function for callback.

```
// Gets the related orders of customer 'foo'.
sharepoint.links({ auth: "...", resource:"Customers", id: "foo", entity: "Order" }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### count(options, callback)

Gets the count of instances in an entity set

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set
* `callback`: A required function for callback.

```
// Gets the total number of customers.
sharepoint.count({ auth: "...", resource:"Customers" }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### create(options, callback)

Adds a new entity instance to an entity set

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set
	* `data`: Object. Required. New entity instance
	* `etag`: String. Optional. ETag value for concurrency control
* `callback`: A required function for callback.

```
var options = { 
	auth: "...", 
	resource: "Customers", 
	data: { /* customer instance /* }
};

// Adds a new customer
sharepoint.create(options, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```


#### replace(options, callback)

Performces a partial update of an existing entity

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set
	* `id`: String or number. Required. Id of the entity to update.
	* `data`: Object. Required. New entity data
	* `etag`: String. Optional. ETag value for concurrency control
* `callback`: A required function for callback.

```
var options = { 
	auth: "...", 
	resource: "Customers",
	id: 10, 
	data: { /* customer instance /* },
	etag: "..."
};

// Updates the customer with id equal to ten.
sharepoint.replace(options, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```


#### update(options, callback)

Performces a complete update of an existing entity

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set
	* `id`: String or number. Required. Id of the entity to update.
	* `data`: Object. Required. New entity data
	* `etag`: String. Optional. ETag value for concurrency control
* `callback`: A required function for callback.

```
var options = { 
	auth: "...", 
	resource: "Customers",
	id: 10, 
	data: { /* customer instance /* },
	etag: "..."
};

// Updates the customer with id equal to ten.
sharepoint.update(options, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```

#### remove(options, callback)

Removes an entity instance from its entity set.

**Parameters:**
* `options`: A required object with he following properties:
	* `auth`: String. Optional. Authentication token 
	* `username`: String. Optional.
	* `password`: String. Optional.
	* `resource`: String. Required. Name of the entity set.
	* `id`: String or number. Required. Id of the entity.
* `callback`: A required function for callback.

```
// Removes the list with id equal to one.
sharepoint.Remove({ auth: "...", resource:"Lists", id:1 }, function(err, result) {
	if (err) return console.error (err);
	console.log (result); 
});
```
