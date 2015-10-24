#tag Class
Protected Class MongoClient
Inherits TCPSocket
	#tag Event
		Sub DataAvailable()
		  Dim dataMB As MemoryBlock
		  
		  timeoutStart = Microseconds
		  
		  while BytesAvailable > 0 
		    
		    dataMB = ReadAll()
		    DataBuffer.Append dataMB
		    
		    if ExpectedMessageSize <= 0 then
		      ' get expected data size from MongoDB message header
		      ExpectedMessageSize = datamb.Int32Value(0)
		      MessageSizeReceived = 0
		    end if
		    
		    MessageSizeReceived = MessageSizeReceived + dataMB.Size
		    
		  wend 
		  
		  if (MessageSizeReceived >= ExpectedMessageSize) then
		    Status = ClientStatus.Idle
		  end if
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Sub connectMongo()
		  Connect()
		  
		  ' wait for connection to be made, with timeout set at 5 seconds
		  
		  timeoutStart = Microseconds
		  do
		    Poll()
		    App.DoEvents()
		  loop until IsConnected or ((Microseconds - timeoutStart) > TCP_TIMEOUT * 1000) or (LastErrorCode <> 0)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  ' call the super constructor
		  
		  Super.Constructor
		  
		  ' configure the default MongoDB connection settings
		  
		  Me.Address = "localhost" 
		  Me.Port = 27017
		  
		  connectMongo ' connect to MongoDB
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1000
		Sub Constructor(initAddress As String)
		  ' call the super constructor
		  
		  Super.Constructor
		  
		  ' configure the given connection settings
		  
		  Me.Address = initAddress
		  Me.Port = 27017
		  
		  connectMongo ' connect to MongoDB
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1000
		Sub Constructor(initAddress As String, initPort As Integer)
		  ' call the super constructor
		  
		  Super.Constructor
		  
		  ' configure the given connection settings
		  
		  Me.Address = initAddress
		  Me.Port = initPort
		  
		  connectMongo ' connect to MongoDB
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Delete(fullCollectionName As String, selector As String, singleRemove As Boolean = True)
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim selectorBSON As MemoryBlock
		  Dim i As Integer
		  Dim flags As Int32
		  
		  ' first convert data into BSON format
		  
		  selectorBSON = BSONSerializer.EncodeJSON(selector)
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + LenB(fullCollectionName) + 1 + 4 + selectorBSON.Size
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_DELETE ' opCode
		  
		  ' OP_DELETE information
		  
		  msg.Int32Value(16) = 0 ' reserved int32
		  
		  msg.CString(20) = fullCollectionName ' fullCollectionName - "dbname.collectionname"
		  pos = 20 + LenB(fullCollectionName) + 1
		  
		  flags = 0
		  if singleRemove then
		    flags = flags + 1 ' set bit 0
		  end if
		  msg.Int32Value(pos) = flags ' update flags
		  pos = pos + 4
		  
		  ' selector - the query to select the document
		  
		  i = 0
		  while i < selectorBSON.Size
		    msg.Byte(pos) = selectorBSON.Byte(i)
		    pos = pos + 1
		    i = i + 1
		  wend
		  
		  ' OP_DELETE does not get  a response from the server, so we simple send the message down our TCP pipe
		  
		  Write msg
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getDB(name As String) As MongoDriver.MongoDatabase
		  Dim db As MongoDriver.MongoDatabase
		  
		  if IsConnected then ' are we connected to Mongo?
		    
		    db = new MongoDriver.MongoDatabase(name, me)
		    
		  end if
		  
		  return db
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getDBNameList() As String()
		  Dim names() As String
		  Dim cursor As MongoDriver.MongoCursor
		  Dim i As Integer
		  Dim j As Integer
		  Dim node As JSONItem
		  Dim dbArrNode As JSONItem
		  Dim dbNode As JSONItem
		  Dim dbName As String
		  
		  cursor = Query("admin.$cmd", "{listDatabases:1}", "", 1)
		  
		  if cursor.DocumentCount > 0 then
		    
		    node = new JSONItem(cursor.Document(0))
		    
		    i = 0 
		    while i < node.Count
		      
		      if node.Name(i) = "databases" then
		        
		        dbArrNode = node.Child("databases")
		        
		        if dbArrNode.IsArray then
		          
		          j = 0
		          while j < dbArrNode.Count
		            
		            dbNode = dbArrNode.Child(j)
		            dbName = dbNode.Lookup("name", "")
		            
		            if dbName <> "" then
		              names.Append dbNode.Lookup("name", "")
		            end if
		            
		            j = j + 1
		          wend
		          
		        end if
		        
		      end if
		      
		      i = i + 1
		    wend
		    
		  end if
		  
		  return names
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GetMore(ByRef cursor As MongoDriver.MongoCursor)
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim i As Integer
		  Dim response As MemoryBlock
		  Dim requestID As Integer
		  Dim responseTo As Integer
		  Dim opCode As Integer
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + LenB(cursor.FullCollectionName) + 1 + 4 + 8
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_GET_MORE ' opCode
		  
		  ' OP_GET_MORE information
		  
		  msg.Int32Value(16) = 0 ' reserved for future use
		  
		  msg.CString(20) = cursor.FullCollectionName ' fullCollectionName - "dbname.collectionname"
		  pos = 20 + LenB(cursor.FullCollectionName) + 1
		  
		  msg.Int32Value(pos) = cursor.NumberToReturn ' numberToReturn - number of documents to return (in the first OP_REPLY batch)
		  pos = pos + 4
		  
		  msg.Int64Value(pos) = cursor.ID ' cursorID - cursorID from the OP_REPLY
		  pos = pos + 8
		  
		  response = sendMessage(msg, true) ' Mongo QUERIES need to wait for responses
		  
		  if response.Size > 32 then
		    
		    response.LittleEndian = true
		    
		    ' parse response MongoDB message
		    
		    msgLen = response.Int32Value(0)
		    requestID = response.Int32Value(4)
		    responseTo = response.Int32Value(8)
		    opCode = response.Int32Value(12)
		    
		    if opCode = OP_REPLY then
		      
		      cursor.ResponseFlags = response.Int32Value(16)
		      cursor.ID = response.Int64Value(20)
		      cursor.NumberToReturn = cursor.NumberToReturn
		      cursor.StartingFrom = response.Int32Value(28)
		      cursor.NumberReturned = response.Int32Value(32)
		      
		      pos = 36
		      
		      for i = 1 to cursor.NumberReturned
		        cursor.Document.Append BSONSerializer.DecodeBSON(response, pos)
		      next i 
		      
		    end if
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Insert(fullCollectionName As String, docs() As String)
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim i As Integer
		  Dim j As Integer
		  'Dim tmpMB As MemoryBlock
		  Dim docsBSON() As MemoryBlock
		  Dim bsonLen As Integer
		  
		  ' first convert documents into BSON format
		  
		  bsonLen = 0
		  for i = 0 to docs.Ubound
		    docsBSON.Append BSONSerializer.EncodeJSON(docs(i))
		    bsonLen = bsonLen +docsBSON(docsBSON.Ubound).Size
		  next i 
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + LenB(fullCollectionName) + 1 + bsonLen
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_INSERT ' opCode
		  
		  ' OP_INSERT information
		  
		  msg.Int32Value(16) = 0 ' flags - bit vector of insert options
		  
		  msg.CString(20) = fullCollectionName ' fullCollectionName - "dbname.collectionname"
		  pos = 20 + LenB(fullCollectionName) + 1
		  
		  for i = 0 to docs.Ubound
		    
		    j = 0
		    while j < docsBSON(i).Size
		      msg.Byte(pos) = docsBSON(i).Byte(j)
		      pos = pos + 1
		      j = j + 1
		    wend
		    
		  next i 
		  
		  ' OP_INSERT does not get  a response from the server, so we simple send the message down our TCP pipe
		  
		  Write msg
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub KillCursors(cursorIDs() As Int64)
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim i As Integer
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + 4 + ((cursorIDs.Ubound + 1) * 8)
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_KILL_CURSORS ' opCode
		  
		  ' OP_KILL_CURSORS information
		  
		  msg.Int32Value(16) = 0 ' reserved for future use
		  
		  msg.Int32Value(20) = (cursorIDs.Ubound + 1) ' numberOfCursorIDs - number of cursorIDs in message
		  
		   ' cursorIDs - “array” of cursor IDs to be closed.
		  
		  pos = 24
		  i = 0
		  while i <= cursorIDs.Ubound
		    msg.Int64Value(pos) = cursorIDs(i)
		    i = i + 1
		    pos = pos + 8
		  wend
		  
		  ' OP_KILL_CURSORS does not get  a response from the server, so we simple send the message down our TCP pipe
		  
		  Write msg
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub KillCursors(cursorID as Int64)
		  Dim cursorIDs() As Int64
		  
		  cursorIDs.Append cursorID
		  KillCursors(cursorIDs)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Query(fullCollectionName As String, query As String, selector As String = "", numberToReturn As Integer = 0) As MongoDriver.MongoCursor
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim queryBSON As MemoryBlock
		  Dim selectorBSON As MemoryBlock
		  Dim i As Integer
		  Dim response As MemoryBlock
		  Dim requestID As Integer
		  Dim responseTo As Integer
		  Dim opCode As Integer
		  Dim cursor As MongoDriver.MongoCursor
		  
		  ' first convert data into BSON format
		  
		  queryBSON = BSONSerializer.EncodeJSON(query)
		  if selector  <> "" then
		    selectorBSON = BSONSerializer.EncodeJSON(selector)
		  else 
		    selectorBSON = ""
		  end if
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + LenB(fullCollectionName) + 1 + 4 + 4 + queryBSON.Size + selectorBSON.Size
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_QUERY ' opCode
		  
		  ' OP_QUERY information
		  
		  msg.Int32Value(16) = 0 ' flags - bit vector of query options
		  
		  msg.CString(20) = fullCollectionName ' fullCollectionName - "dbname.collectionname"
		  pos = 20 + LenB(fullCollectionName) + 1
		  
		  msg.Int32Value(pos) = 0 ' numberToSkip - number of documents to skip
		  pos = pos + 4
		  
		  msg.Int32Value(pos) = numberToReturn ' numberToReturn - number of documents to return (in the first OP_REPLY batch)
		  pos = pos + 4
		  
		  ' query - BSON document that represents the query
		  
		  i = 0 
		  while i < queryBSON.Size
		    msg.Byte(pos) = queryBSON.Byte(i)
		    pos = pos + 1
		    i = i + 1
		  wend
		  
		  ' selector - BSON document that indicates the fields to return
		  
		  i = 0
		  while i < selectorBSON.Size
		    msg.Byte(pos) = selectorBSON.Byte(i)
		    pos = pos + 1
		    i = i + 1
		  wend
		  
		  response = sendMessage(msg, true) ' Mongo QUERIES need to wait for responses
		  
		  if response.Size > 32 then
		    
		    response.LittleEndian = true
		    
		    ' parse response MongoDB message
		    
		    msgLen = response.Int32Value(0)
		    requestID = response.Int32Value(4)
		    responseTo = response.Int32Value(8)
		    opCode = response.Int32Value(12)
		    
		    if opCode = OP_REPLY then
		      
		      cursor = new MongoCursor(me)
		      
		      cursor.FullCollectionName = fullCollectionName
		      
		      cursor.ResponseFlags = response.Int32Value(16)
		      cursor.ID = response.Int64Value(20)
		      cursor.NumberToReturn = numberToReturn
		      cursor.StartingFrom = response.Int32Value(28)
		      cursor.NumberReturned = response.Int32Value(32)
		      
		      pos = 36
		      
		      for i = 1 to cursor.NumberReturned
		        cursor.Document.Append BSONSerializer.DecodeBSON(response, pos)
		      next i 
		      
		    end if
		    
		  end if
		  
		  return cursor
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function sendMessage(msg As MemoryBlock, waitForData As Boolean = false) As String
		  Dim response As String
		  
		  if waitForData then
		    Status = ClientStatus.AwaitingResponse
		  else
		    Status = ClientStatus.Idle
		  end if
		  
		  ExpectedMessageSize = 0
		  MessageSizeReceived = 0
		  Write msg
		  
		  timeoutStart = Microseconds
		  do
		    ' wait for data to arrive
		    App.DoEvents()
		  loop until (Status = ClientStatus.Idle) or  ((Microseconds - timeoutStart) > TCP_TIMEOUT * 1000)
		  
		  response = Join(DataBuffer, "")
		  Redim DataBuffer(-1)
		  ExpectedMessageSize = 0
		  MessageSizeReceived = 0
		  
		  return response
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Update(fullCollectionName As String, selector As String, update As String = "", upsert As Boolean = False, multiUpdate As Boolean = False)
		  Dim msg As MemoryBlock
		  Dim msgLen As Int32
		  Dim pos As Integer
		  Dim selectorBSON As MemoryBlock
		  Dim updateBSON As MemoryBlock
		  Dim i As Integer
		  Dim flags As Int32
		  
		  ' first convert data into BSON format
		  
		  selectorBSON = BSONSerializer.EncodeJSON(selector)
		  updateBSON = BSONSerializer.EncodeJSON(update)
		  
		  ' create memory block for message
		  
		  msgLen = 16 + 4 + LenB(fullCollectionName) + 1 + 4 + selectorBSON.Size + updateBSON.Size
		  msg = new MemoryBlock(msgLen)
		  msg.LittleEndian = true
		  
		  ' standard message header
		  
		  msg.Int32Value(0) = msgLen ' messageLength - total message size, including this
		  msg.Int32Value(4) = 0 ' requestID - identifier for this message
		  msg.Int32Value(8) = 0 ' responseTo - requestID from the original request (used in reponses from db)
		  msg.Int32Value(12) = MongoDriver.OP_UPDATE ' opCode
		  
		  ' OP_UPDATE information
		  
		  msg.Int32Value(16) = 0 ' reserved int32
		  
		  msg.CString(20) = fullCollectionName ' fullCollectionName - "dbname.collectionname"
		  pos = 20 + LenB(fullCollectionName) + 1
		  
		  flags = 0
		  if upsert then
		    flags = flags + 1 ' set bit 0
		  end if
		  if multiUpdate then
		    flags = flags + 2 ' set bit 1
		  end if
		  msg.Int32Value(pos) = flags ' update flags
		  pos = pos + 4
		  
		  ' selector - the query to select the document
		  
		  i = 0
		  while i < selectorBSON.Size
		    msg.Byte(pos) = selectorBSON.Byte(i)
		    pos = pos + 1
		    i = i + 1
		  wend
		  
		  ' update - specification of the update to perform
		  
		  i = 0
		  while i < updateBSON.Size
		    msg.Byte(pos) = updateBSON.Byte(i)
		    pos = pos + 1
		    i = i + 1
		  wend
		  
		  ' OP_UPDATE does not get  a response from the server, so we simple send the message down our TCP pipe
		  
		  Write msg
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected DataBuffer() As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected ExpectedMessageSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected MessageSizeReceived As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Status As ClientStatus
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected timeoutStart As Double
	#tag EndProperty


	#tag Enum, Name = ClientStatus, Type = Integer, Flags = &h21
		Idle
		AwaitingResponse
	#tag EndEnum


	#tag ViewBehavior
		#tag ViewProperty
			Name="Address"
			Visible=true
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType="Integer"
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
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Port"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
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
