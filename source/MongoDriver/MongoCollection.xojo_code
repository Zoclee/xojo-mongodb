#tag Class
Protected Class MongoCollection
	#tag Method, Flags = &h0
		Sub Constructor(initName As String, initDatabase As MongoDriver.MongoDatabase)
		  ' initialize the collection
		  
		  mName = initName
		  mDatabase = initDatabase
		  mClient = mDatabase.Client
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function copyTo(newCollection As String) As Integer
		  Dim count As Integer
		  Dim toColl As MongoDriver.MongoCollection
		  Dim cursor As MongoDriver.MongoCursor
		  Dim doc As String
		  
		  toColl = mDatabase.getCollection(newCollection)
		  call toColl.ensureIndex("{_id:1}")
		  
		  count = 0
		  cursor = find()
		  while cursor.hasNext()
		    doc = cursor.getNext()
		    count = count + 1
		    toColl.save(doc)
		  wend
		  
		  return count
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function count(query As String = "") As Integer
		  Dim result As String
		  Dim count As Integer
		  Dim resultJSON As JSONItem
		  Dim cmd As String
		  
		  if query <> "" then
		    cmd = "{count:""" + mName + """,query:" + query + "}"
		  else
		    cmd = "{count:""" + mName + """}"
		  end if
		  
		  result = mDatabase.runCommand(cmd)
		  resultJSON = new JSONItem(result)
		  count = resultJSON.Lookup("n", 0)
		  
		  return count
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function dataSize() As Integer
		  Dim statsDoc As String
		  Dim statsJSON As JSONItem
		  Dim result As Integer 
		  
		  result = 0
		  
		  statsDoc = stats()
		  statsJSON = ParseJSON(statsDoc)
		  
		  result = statsJSON.Lookup("size", 0)
		  
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function distinct(field As String, query As String = "") As String
		  Dim result As String
		  dim resultJSON As JSONItem
		  
		  if query <> "" then
		    result = mDatabase.runCommand("{distinct:""" + mName + """,key:""" + field + """,query:" + query + "}")
		  else
		    result = mDatabase.runCommand("{distinct:""" + mName + """,key:""" + field + """,query:{}}")
		  end if
		  
		  resultJSON = ParseJSON(result)
		  resultJSON = resultJSON.Value("values")
		  result = resultJSON.ToString()
		  
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function drop() As Boolean
		  Dim result As String
		  Dim resultJSON As JSONItem
		  Dim cmd As String
		  Dim cmdResult As Boolean
		  
		  cmd = "{drop:""" + mName + """}"
		  
		  result = mDatabase.runCommand(cmd)
		  resultJSON = new JSONItem(result)
		  
		  if resultJSON.Value("ok") = 1 then
		    cmdResult = true
		  else
		    cmdResult = false
		  end if
		  
		  return cmdResult
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function dropIndex(index As String) As String
		  Dim result As String
		  
		  if IsJSONObject(index) then
		    result = mDatabase.runCommand("{deleteIndexes:""" + mName + """,index:" + index + "}")
		  else
		    result = mDatabase.runCommand("{deleteIndexes:""" + mName + """,index:""" + index + """}")
		  end if
		  
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function dropIndexes() As String
		  Dim result As String
		  
		  result = mDatabase.runCommand("{deleteIndexes:""" + mName + """,index:""*""}")
		  
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ensureIndex(keys As String, unique As Boolean) As String
		  Dim uniqueStr As String
		  
		  if unique then
		    uniqueStr = "true"
		  else
		    uniqueStr = "false"
		  end if
		  
		  return ensureIndex(keys, "{unique:" + uniqueStr + "}")
		   
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ensureIndex(keys As String, unique As Boolean, dropDups As Boolean) As String
		  Dim uniqueStr As String
		  Dim dropDupsStr As String
		  
		  if unique then
		    uniqueStr = "true"
		  else
		    uniqueStr = "false"
		  end if
		  
		  if dropDups then
		    dropDupsStr = "true"
		  else
		    dropDupsStr = "false"
		  end if
		  
		  return ensureIndex(keys, "{unique:" + uniqueStr + ",dropDups:" + dropDupsStr + "}")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ensureIndex(keys As String, unique As Boolean, dropDups As Boolean, background As Boolean) As String
		  Dim uniqueStr As String
		  Dim dropDupsStr As String
		  Dim backgroundStr As String
		  
		  if unique then
		    uniqueStr = "true"
		  else
		    uniqueStr = "false"
		  end if
		  
		  if dropDups then
		    dropDupsStr = "true"
		  else
		    dropDupsStr = "false"
		  end if
		  
		  if background then
		    backgroundStr = "true"
		  else
		    backgroundStr = "false"
		  end if
		  
		  return ensureIndex(keys, "{unique:" + uniqueStr + ",dropDups:" + dropDupsStr + ",background:" + backgroundStr + "}")
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ensureIndex(keys As String, options As String = "") As String
		  Dim fullName As String
		  Dim indexName As String
		  Dim jsonData() As String
		  Dim indexDoc As String
		  Dim coll As MongoDriver.MongoCollection
		  Dim prepOptions As String
		  Dim err As String
		  
		  err = ""
		  
		  if mClient.IsConnected then
		    
		    fullName = mDatabase.getName() + "." + mName
		    indexName = genIndexName(keys) 
		    
		    jsonData.Append "{ns:"""
		    jsonData.Append fullName
		    jsonData.Append """,key:"
		    jsonData.Append keys
		    jsonData.Append ",name:"""
		    jsonData.Append indexName
		    jsonData.Append """"
		    
		    // add options
		    
		    if Trim(options) <> "" then
		      prepOptions = Trim(options)
		      prepOptions = Mid(prepOptions, 2, Len(prepOptions) - 2) ' remove start end end curly brackets
		      jsonData.Append ","
		      jsonData.Append prepOptions
		    end if
		    
		    jsonData.Append "}"
		    indexDoc = join(jsonData, "")
		    
		    coll = mDatabase.getCollection("system.indexes")
		    coll.insert(indexDoc)
		    
		    err = mDatabase.getLastError
		    
		  end if
		  
		  return err
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function find(criteria As String = "", projection As String = "") As MongoDriver.MongoCursor
		  Dim cursor As MongoDriver.MongoCursor
		  
		  cursor = mClient.Query(mDatabase.getName() + "." + mName, criteria, projection)
		  
		  return cursor
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function findAndModify(document As String) As String
		  Dim result As String
		  Dim cmd As String
		  Dim resultJSON As JSONItem
		  
		  cmd = "{findandmodify:""" + mName + """," + StripJSONBrackets(document) + "}"
		  
		  result = mDatabase.runCommand(cmd)
		  resultJSON = ParseJSON(result)
		  if resultJSON.Value("ok") = 1 then
		    resultJSON = resultJSON.Value("value")
		    result = resultJSON.ToString()
		  else
		    result = ""
		  end if
		  
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function findOne(criteria As String = "", projection As String = "") As String
		  Dim cursor As MongoDriver.MongoCursor
		  
		  cursor = mClient.Query(mDatabase.getName() + "." + mName, criteria, projection, 1, me)
		  
		  if cursor <> nil then
		    return cursor.getNext()
		  else
		    return ""
		  end if
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function genIndexName(keys As String) As String
		  Dim name As String
		  Dim json As JSONItem
		  Dim i As Integer
		  
		  json = JSONSerializer.ParseJSON(keys)
		  
		  name = ""
		  
		  i = 0
		  while i < json.Count
		    if Len(name) > 0 then
		      name = name + "_"
		    end if
		    name = name + json.Name(i) + "_"
		    name = name + json.Lookup(json.Name(i), "")
		    i = i + 1
		  wend
		  
		  return name
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getIndexes() As String()
		  Dim result() As String
		  Dim coll As MongoDriver.MongoCollection
		  Dim fullName As String
		  Dim cursor As MongoDriver.MongoCursor
		  
		  fullName = mDatabase.getName() + "." + mName
		  
		  coll = mDatabase.getCollection("system.indexes")
		  
		  cursor = coll.find("{ns:""" + fullName + """}")
		  
		  while cursor.hasNext
		    result.Append cursor.getNext()
		  wend
		  
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub insert(documents() As String)
		  mClient.Insert mDatabase.getName() + "." + mName, documents
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub insert(document As String)
		  Dim doc() As String
		  
		  doc.Append document
		  insert(doc)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function reIndex() As String
		  Dim result As String
		  
		  result = mDatabase.runCommand("{reIndex:""" + mName + """}")
		  
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub remove(query As String, justOne As Boolean = true)
		  mClient.Delete mDatabase.getName() + "." + mName, query, justOne
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub save(document As String)
		  Dim docJSON As JSONItem
		  Dim id As Variant
		  Dim idItem As JSONItem
		  
		  docJSON = ParseJSON(document)
		  
		  id = docJSON.Lookup("_id", nil)
		  
		  if id = nil then
		    
		    ' insert object
		    
		    insert(document)
		    
		  else
		    
		    ' update object
		    
		    idItem=  new JSONItem
		    idItem.Value("_id") = id
		    
		    update(idItem.ToString(), document, true)
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function stats(scale As Integer = 1) As String
		  Dim result As String
		  
		  result = mDatabase.runCommand("{collstats:""" + mName + """,scale:"  + Str(scale) + "}")
		   
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function totalIndexSize() As Integer
		  Dim result As Integer
		  Dim statsJSON As JSONItem
		  
		  statsJSON = JSONSerializer.ParseJSON(stats())
		  result=  statsJSON.Lookup("totalIndexSize", 0)
		  
		  Return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub update(query As String, update As String, upsert As Boolean = false, multi As Boolean = false)
		  mClient.Update mDatabase.getName() + "." + mName, query, update, upsert, multi
		  
		End Sub
	#tag EndMethod


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mDatabase
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mDatabase = value
			End Set
		#tag EndSetter
		Database As MongoDriver.MongoDatabase
	#tag EndComputedProperty

	#tag Property, Flags = &h1
		Protected mClient As MongoDriver.MongoClient
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mDatabase As MongoDriver.MongoDatabase
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mName As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
