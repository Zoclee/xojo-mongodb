#tag Class
Protected Class MongoDatabase
	#tag Method, Flags = &h0
		Function auth(username As String, password As String) As String
		  Dim result As String
		  Dim tmpObj As JSONItem
		  Dim nonce As String
		  Dim digest As String
		  Dim pwdDigest As String
		  Dim authCmd As String
		  
		  //  get a nonce for use in authentication
		  
		  tmpObj = new JSONItem(runCommand("{ getnonce: 1}"))
		  nonce = tmpObj.Lookup("nonce", "")
		  
		  // compile authentication digest
		  
		  pwdDigest = Lowercase(EncodeHex(MD5(username + ":mongo:" + password)))
		  digest = Lowercase(EncodeHex(MD5(nonce + username + pwdDigest)))
		  
		  //  authenticate
		  
		  authCmd = "{authenticate : 1,user :""" + username + """,nonce :""" + nonce + """,key :""" + digest + """}"
		  
		  result = runCommand(authCmd)
		  
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(initName As String, initClient As MongoDriver.MongoClient)
		  ' initialize the database
		  
		  mName = initName
		  mClient = initClient
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getCollection(name As String) As MongoDriver.MongoCollection
		  Dim coll As MongoDriver.MongoCollection
		  
		  if mClient.IsConnected then ' are we connected to Mongo?
		    
		    coll = new MongoDriver.MongoCollection(name, mClient, me)
		    
		  end if
		  
		  return coll
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getCollectionNames() As String()
		  Dim names() As String
		  Dim cursor As MongoDriver.MongoCursor
		  Dim node As JSONItem
		  Dim colName As String
		  Dim dbNameLen As Integer
		  Dim i As Integer
		  
		  ' todo: use $regex to minimize results
		  cursor = mClient.Query(mName + ".system.namespaces", "{$query:{},$orderby:{name:1}}", "{name:1}")
		  
		  dbNameLen = Len(mName) + 1
		  
		  i = 0
		  while i < cursor.DocumentCount
		    
		    node = new JSONItem(cursor.Document(i))
		    
		    colName = node.Lookup("name", "")
		    
		    if (colName <> "") and (Right(colName, 6) <> ".$_id_") then
		      colName = Right(colName, Len(colName) - dbNameLen)
		      names.Append colName
		    end if
		    
		    i = i + 1
		    
		  wend
		  
		  return names
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getLastError() As String
		  Dim errDoc As String
		  Dim errJSON As JSONItem
		  Dim err As String
		  
		  err = ""
		  
		  errDoc = getLastErrorObj()
		  if errDoc <>"" then
		    errJSON = ParseJSON(errDoc)
		    err = errJSON.Lookup("err", "")
		  end if
		  
		  return err
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getLastErrorObj() As String
		  Dim errDoc As String
		  Dim cursor As MongoDriver.MongoCursor
		  
		  errDoc = ""
		  
		  if mClient.IsConnected then ' are we connected to Mongo?
		    
		    cursor = mClient.Query("admin.$cmd", "{getLastError:1}", "", 1)
		    if cursor.hasNext then
		      errDoc = cursor.getNext
		    end if
		    
		  end if
		  
		  return errDoc
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getName() As string
		  return mName
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function runCommand(command As String) As String
		  Dim coll As MongoDriver.MongoCollection
		  Dim result As String
		  
		  coll = getCollection("$cmd")
		  
		  if IsJSONObject(command) then
		    result = coll.findOne(command)
		  else
		    result = coll.findOne("{" + command + ":1}")
		  end if
		  
		  return result
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mClient As MongoDriver.MongoClient
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
