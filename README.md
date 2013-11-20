MongoDB Xojo Driver
===================

This is the MongoDB driver for the Xojo programming language.

The current version of the driver only supports connections to a single server. Support for replica sets and sharding will be added in future versions of the driver.

Installation
------------

Simply import the BSONSerializer, JSONSerializer and MongoDriver modules into your project. The module files are located in the [src folder](https://github.com/alwyn1024/mongodb-xojodriver/tree/master/src) of the repository.

Example Code
------------

	Dim client As new MongoDriver.MongoClient() ' connect to localhost on default port
	Dim db As MongoDriver.MongoDatabase
	Dim coll As MongoDriver.MongoCollection
	Dim cursor As MongoDriver.MongoCursor
	Dim result As String
	Dim resultArr () As String
	Dim resultInt As Integer
	
	if client.IsConnected then
	
		db = client.getDB("test") ' get database object
		
		coll = db.getCollection("data") ' get collection object
		
		resultInt = coll.count() ' get number of documents in collection
		
		coll.insert "{mydoc:123}" ' insert document into collection
		
		result = coll.findAndModify("{query:{mydoc:123},update:{mydoc:456}}") ' modify and return a single document
		
		' query and loop through documents
		
		cursor = coll.find("{}")
		while cursor.hasNext
			result = cursor.getNext()
		wend
		
		coll.remove("{a:1}") ' remove document
		
		coll.update"{mydoc:123}", "{$set:{mydoc:234}}" ' update document
		
		result = db.getLastErrorObj() ' get last error JSON
		
		result = db.getLastError() ' get last error message
		
		result = db.runCommand("{count:""test:data""}") ' run database command
		
		result = coll.ensureIndex("{myItem:1}") ' create ascending index on field myItem
		
		resultArr = coll.getIndexes() ' get list of indexes
		
		result = coll.dropIndex("{myItem:1}") ' remove index from collection
		
		result = coll.dropIndexes() ' remove all indexes from collection
		
		result = coll.stats(1024) ' get collection stats scaled to kilobyte values
		
		coll.save"{_id:1,test:""abc""}" ' insert document
		coll.save"{_id:1,test:""def""}" ' update document
		
		resultInt = coll.totalIndexSize() ' get total size of all indexes for the collection
		
		resultInt = coll.copyTo("dataCopy") ' copy whole collection
		
		resultInt = coll.dataSize() ' get the size of the collection in bytes
		
		result = coll.distinct("mydoc") ' find distinct values for a specified field
		
		'coll = db.getCollection("somecoll") ' select collection to remove
		'resultBool = coll.drop() ' remove collection from database
		
		result = coll.reIndex() ' drop and recreate all indexes on collection

	end if