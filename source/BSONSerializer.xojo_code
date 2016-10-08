#tag Module
Protected Module BSONSerializer
	#tag Method, Flags = &h21
		Private Function bsonStr(d As Double) As String
		  Dim s As String
		  
		  if (d \ 1) = d then // whole number?
		    s = Str(d, "-0")
		  else
		    s = Str(d, "-0.#######")
		    if Right(s, 1) = "." then
		      s = Str(d, "-0")
		    end if
		  end if
		  
		  return s
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function decExtractCString(bsonMB As MemoryBlock, ByRef pos As Integer) As String
		  Dim startPos As Integer = pos
		  
		  while bsonMB.Byte(pos) <> 0
		    pos = pos + 1
		  wend
		  
		  pos = pos + 1
		  
		  Dim strSize As Integer = pos - startPos - 1
		  
		  If strSize > 0 Then
		    Dim s As String = DefineEncoding(bsonMB.StringValue(startPos, strSize), Encodings.UTF8)
		    return escapeJSON(s)
		  End If
		  
		  return ""
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DecodeBSON(bsonMB As MemoryBlock, ByRef pos As Integer) As String
		  Dim json As String
		  
		  json = decodeDocument(bsonMB, pos)
		  
		  return json
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DecodeBSON(bson As String) As String
		  Dim bsonMB As MemoryBlock = bson
		  Dim pos As Integer = 0
		  Dim json As String
		  
		  json = decodeDocument(bsonMB, pos)
		  
		  return json
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function decodeDocument(bsonMB As MemoryBlock, ByRef pos As Integer, parseAsArray As Boolean = False) As String
		  Dim json() As String
		  Dim pairName As String
		  Dim pairValue As String
		  Dim firstPair As Boolean
		  Dim i As Integer
		  Dim objIdValue() As String
		  
		  if parseAsArray then
		    json.Append "["
		  else
		    json.Append "{"
		  end if
		  
		  firstPair = true
		  
		  pos = pos + 4 ' skip past document size
		  
		  while bsonMB.Byte(pos) <> 0
		    
		    if not firstPair then
		      json.Append ","
		    end if
		    
		    firstPair = false
		    
		    select case bsonMB.Byte(pos)
		      
		    case 1 ' double
		      
		      pos = pos + 1
		      pairName = decExtractCString(bsonMB, pos)
		      
		      pairValue = bsonStr(bsonMB.DoubleValue(pos))
		      pos = pos + 8
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		      
		    case 2 ' utf-8 string
		      
		      pos = pos + 1
		      pairName = decExtractCString(bsonMB, pos)
		      pos = pos + 4 ' skip past string length
		      
		      pairValue = decExtractCString(bsonMB, pos)
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"""
		      else
		        json.Append """"
		      end if
		      json.Append pairValue
		      json.Append """"
		      
		    case 3 ' document
		      
		      pos = pos + 1
		      
		      pairName = decExtractCString(bsonMB, pos)
		      pairValue = decodeDocument(bsonMB, pos)
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		    case 4 ' array
		      
		      pos = pos + 1
		      
		      pairName = decExtractCString(bsonMB, pos)
		      pairValue = decodeDocument(bsonMB, pos, True)
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		      
		    case 7 ' ObjectId
		      
		      pos = pos + 1
		      
		      pairName = decExtractCString(bsonMB, pos)
		      
		      objIdValue.Append """"
		      for i = 1 to 12
		        if bsonMB.Byte(pos) < 16 then
		          objIdValue.Append "0" + Lowercase(Hex(bsonMB.Byte(pos)))
		        else
		          objIdValue.Append Lowercase(Hex(bsonMB.Byte(pos)))
		        end if
		        pos = pos + 1
		      next i
		      objIdValue.Append """"
		      
		      json.Append """"
		      json.Append pairName
		      json.Append """:"
		      json.Append join(objIdValue, "")
		      
		    case 8 ' boolean
		      
		      pos = pos + 1
		      
		      pairName = decExtractCString(bsonMB, pos)
		      
		      if bsonMB.Byte(pos) = 0 then
		        pairValue = "false"
		      else
		        pairValue = "true"
		      end if
		      pos = pos + 1
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		    case 9 ' UTC datetime
		      
		      pos = pos + 1
		      
		      pairName = decExtractCString(bsonMB, pos)
		      
		      pairValue = formatDate(bsonMB.Int64Value(pos))
		      pos = pos + 8
		      
		      json.Append """"
		      json.Append pairName
		      json.Append """:"""
		      json.Append pairValue
		      json.Append """"
		      
		    case 10 ' null
		      
		      pos = pos + 1
		      pairName = decExtractCString(bsonMB, pos)
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:null"
		      else
		        json.Append "null"
		      end if
		      
		    case 16 ' 32-bit integer
		      
		      pos = pos + 1
		      pairName = decExtractCString(bsonMB, pos)
		      
		      pairValue = Str(bsonMB.Int32Value(pos))
		      pos = pos + 4
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		    case 18 ' 64-bit integer
		      
		      pos = pos + 1
		      pairName = decExtractCString(bsonMB, pos)
		      
		      pairValue = Str(bsonMB.Int64Value(pos))
		      pos = pos + 8
		      
		      if not parseAsArray then
		        json.Append """"
		        json.Append pairName
		        json.Append """:"
		      end if
		      json.Append pairValue
		      
		    case else
		      
		      pos = pos + 1 ' to prevent infinite loop
		      break  ' todo
		      
		    end select
		    
		  wend
		  
		  pos = pos + 1
		  
		  if parseAsArray then
		    json.Append "]"
		  else
		    json.Append "}"
		  end if
		  
		  return join(json, "")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function encExtractString(jsonMB As MemoryBlock, ByRef pos As Integer, quoteChar As Byte) As String
		  Dim startPos As Integer
		  Dim sMB As MemoryBlock
		  Dim i As Integer
		  
		  skipWhitespace jsonMB, pos
		  
		  pos = pos + 1
		  
		  startPos = pos
		  
		  while jsonMB.Byte(pos) <> quoteChar ' "
		    pos = pos + 1
		  wend
		  
		  sMB = new MemoryBlock(pos - startPos)
		  pos = pos - 1
		  for i = startPos to pos
		    sMB.Byte(i - startPos) = jsonMB.Byte(i)
		  next i 
		  
		  pos = pos + 2
		  
		  return  sMB
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function encExtractToken(jsonMB As MemoryBlock, ByRef pos As Integer) As String
		  Dim strChars() As String
		  
		  skipWhitespace jsonMB, pos
		  
		  while (jsonMB.Byte(pos) > 32) and (jsonMB.Byte(pos) <> 58) _ ' :
		    and (jsonMB.Byte(pos) <> 125) _ ' }
		    and (jsonMB.Byte(pos) <> 44) _ ' ,
		    and (jsonMB.Byte(pos) <> 93) ' ]
		    
		    strChars.Append Chr(jsonMB.Byte(pos))
		    pos = pos + 1
		    
		  wend
		  
		  return Join(strChars, "")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function EncodeJSON(json As String) As String
		  Dim pos As Integer
		  Dim bson As String
		  
		  pos = 0
		  bson = encodeObject(json, pos)
		  
		  return bson
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function encodeObject(jsonMB As MemoryBlock, ByRef pos As Integer) As String
		  Dim objValues() As String
		  Dim objLen As Integer
		  Dim i As Integer
		  Dim bson As String
		  
		  skipWhitespace jsonMB, pos
		  
		  if pos < jsonMB.Size then
		    
		    if jsonMB.Byte(pos) = 123 then ' {
		      
		      pos = pos + 1
		      skipWhitespace jsonMB, pos
		      
		      if jsonMB.Byte(pos) = 125 then ' }
		        
		        // do nothing (empty JSON object)
		        
		      else
		        
		        while jsonMB.Byte(pos) <> 125 ' }
		          
		          objValues.Append encodePair(jsonMB, pos)
		          
		          skipWhitespace jsonMB, pos
		          
		          if  jsonMB.Byte(pos) = 44 then ' ,
		            pos = pos + 1
		          end if
		          
		          skipWhitespace jsonMB, pos
		          
		        wend
		        
		      end if
		      
		      pos = pos + 1
		      
		    end if
		    
		  end if
		  
		  objLen = 0
		  for i = 0 to objValues.Ubound
		    objLen = objLen + Len(objValues(i))
		  next i
		  objLen = objLen + 5
		  
		  if use64BitInteger(objLen) then
		    bson = formatInt64(objLen) + Join(objValues,"") + Chr(0)
		  else
		    bson = formatInt32(objLen) + Join(objValues, "") + Chr(0)
		  end if
		  
		  return bson
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function encodePair(jsonMB As MemoryBlock, ByRef pos As Integer, name As String = "") As String
		  Dim pairName As String
		  Dim pairValue As String
		  Dim bson() As String
		  Dim arrIndex As Integer
		  Dim arrValues() As String
		  Dim arrLen As Integer
		  Dim i As Integer
		  
		  skipWhitespace jsonMB, pos
		  
		  ' parse name (if not already given)
		  
		  if name = "" then
		    
		    if (jsonMB.Byte(pos) = 34) or (jsonMB.Byte(pos) = 39) then ' " OR '
		      pairName = encExtractString(jsonMB, pos, jsonMB.Byte(pos)) ' quoted name
		    else
		      pairName = encExtractToken(jsonMB, pos) ' non-quoted name
		    end if
		    
		    skipWhitespace jsonMB, pos
		    
		    pos = pos + 1 ' skip past :
		    
		    skipWhitespace jsonMB, pos
		    
		  else
		    
		    pairName = name
		    
		  end if
		  
		  ' parse value
		  
		  select case jsonMB.Byte(pos)
		    
		  case 34, 39 ' " OR ' = string
		    
		    pairValue = encExtractString(jsonMB, pos, jsonMB.Byte(pos))
		    
		    if (pairName = "_id") and isObjectId(pairValue) then
		      bson.Append Chr(7)
		      bson.Append pairName
		      bson.Append Chr(0)
		      bson.Append formatObjectId(pairValue)
		    else
		      bson.Append Chr(2)
		      bson.Append pairName
		      bson.Append Chr(0)
		      bson.Append formatInt32(Len(pairValue) + 1)
		      bson.Append pairValue
		      bson.Append Chr(0)
		    end if
		    
		  case 91 ' [ = array
		    
		    pos = pos + 1
		    
		    skipWhitespace jsonMB, pos
		    
		    bson.Append Chr(4)
		    bson.Append pairName
		    bson.Append Chr(0)
		    
		    arrIndex = -1
		    
		    while jsonMB.Byte(pos) <> 93 ' ]
		      
		      arrIndex = arrIndex + 1
		      
		      arrValues.Append encodePair(jsonMB, pos, str(arrIndex))
		      
		      skipWhitespace jsonMB, pos
		      
		      if  jsonMB.Byte(pos) = 44 then ' ,
		        pos = pos + 1
		        skipWhitespace jsonMB, pos
		      end if
		      
		    wend
		    pos = pos + 1 ' skip past ] 
		    
		    arrLen = 0
		    for i = 0 to arrValues.Ubound
		      arrLen = arrLen + Len(arrValues(i))
		    next i
		    arrLen = arrLen + 5
		    bson.Append formatInt32(arrLen) + Join(arrValues, "") + Chr(0)
		    
		  case 123 ' { = object
		    
		    pairValue = encodeObject(jsonMB, pos)
		    
		    bson.Append Chr(3)
		    bson.Append pairName
		    bson.Append Chr(0)
		    bson.Append pairValue
		    
		  case else
		    
		    pairValue = encExtractToken(jsonMB, pos)
		    
		    if pairValue = "true" then
		      
		      bson.Append Chr(8)
		      bson.Append pairName
		      bson.Append Chr(0)
		      bson.Append Chr(1)
		      
		    elseif pairValue = "false" then
		      
		      bson.Append Chr(8)
		      bson.Append pairName
		      bson.Append Chr(0)
		      bson.Append Chr(0)
		      
		    elseif pairValue = "null" then
		      
		      bson.Append Chr(10)
		      bson.Append pairName
		      bson.Append Chr(0)
		      
		    else
		      
		      ' numeric value
		      
		      if InStr(1, pairValue, ".") > 0 then
		        
		        ' double
		        
		        bson.Append Chr(1)
		        bson.Append pairName
		        bson.Append Chr(0)
		        bson.Append formatDouble(Val(pairValue))
		        
		      else
		        
		        ' int32
		        
		        bson.Append Chr(16)
		        bson.Append pairName
		        bson.Append Chr(0)
		        bson.Append formatInt32(Val(pairValue))
		        
		        ' todo: support for int64
		        
		      end if
		      
		    end if
		    
		  end select
		  
		  return Join(bson, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function escapeJSON(s As String) As String
		  Dim result As String
		  
		  result = ReplaceAll(s, "\", "\\")
		  result = ReplaceAll(result, """", "\""")
		  'result = ReplaceAll(result, "/", "\/") ' this escape code is technically not required
		  result = ReplaceAll(result, chr(8), "\b")
		  result = ReplaceAll(result, chr(12), "\f")
		  result = ReplaceAll(result, chr(10), "\n")
		  result = ReplaceAll(result, chr(13), "\r")
		  result = ReplaceAll(result, chr(9), "\t")
		  
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function formatDate(millseconds As Int64) As String
		  Dim result() As String
		  Dim d As new Date
		  Dim frac As Integer
		  
		  d.TotalSeconds = (millseconds \ 1000) + EPOCH_SECONDS
		  
		  // 2013-03-16T02:50:27.874Z
		  
		  result.Append Str(d.Year) 
		  result.Append "-"
		  result.Append Str(d.Month, "00")
		  result.Append "-"
		  result.Append Str(d.Day, "00")
		  result.Append "T"
		  result.Append Str(d.Hour, "00")
		  result.Append ":"
		  result.Append Str(d.Minute, "00")
		  result.Append ":"
		  result.Append Str(d.Second, "00")
		  
		  frac = millseconds mod 1000
		  if frac <> 0 then
		    result.Append "."
		    result.Append Str(frac)
		  end if
		  
		  result.Append "Z" ' to indicate UTC
		  
		  return Join(result, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function formatDouble(d As Double) As String
		  Dim tmpMB As new MemoryBlock(8)
		  
		  tmpMB.LittleEndian = true
		  tmpMB.DoubleValue(0) = d
		  
		  Return tmpMB
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function formatInt32(i As Int32) As String
		  Dim tmpMB As new MemoryBlock(4)
		  
		  tmpMB.LittleEndian = true
		  tmpMB.Int32Value(0) = i
		  
		  Return tmpMB
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function formatInt64(i As Int64) As String
		  Dim tmpMB As new MemoryBlock(8)
		  
		  tmpMB.LittleEndian = true
		  tmpMB.Int64Value(0) = i
		  
		  Return tmpMB
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function formatObjectId(s As String) As String
		  dim tmpMB As new MemoryBlock(12)
		  dim i, j As Integer
		  
		  j = 1
		  
		  for i = 0 to 11
		    tmpMB.Byte(i) = val("&h" + s.Mid(j,2))
		    j = j + 2
		  next
		  
		  return tmpMB
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function isObjectId(s As String) As Boolean
		  dim i, c, l As Integer
		  
		  l = s.LenB
		  
		  if l <> 24 then Return False
		  
		  for i = 1 to l
		    c = s.Mid(i, 1).Uppercase.Asc
		    if not (( c >= 48 and c <= 57 ) or ( c >= 65 and c <= 70)) then Return False
		  next
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub skipWhitespace(jsonMB As MemoryBlock, ByRef pos As Integer)
		  Dim bDone As Boolean
		  
		  if pos < jsonMB.Size then
		    
		    bDone = false
		    
		    do
		      if jsonMB.Byte(pos) <= 32 then
		        pos = pos + 1
		        if pos >= jsonMB.Size then
		          bDone = true
		        end if
		      else
		        bDone = true
		      end if
		    loop until bDone
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function use64BitInteger(param As integer) As Boolean
		  Dim result As Boolean
		  
		  result = false
		  
		  if (param < BSONSerializer.kMin32BitIntegerValue) or (param > BSONSerializer.kMax32BitIntegerValue) then
		    result = true
		  end if
		  
		  return result
		  
		End Function
	#tag EndMethod


	#tag Constant, Name = EPOCH_SECONDS, Type = Double, Dynamic = False, Default = \"2082844800", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kMax32BitIntegerValue, Type = Double, Dynamic = False, Default = \"2147483647", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kMin32BitIntegerValue, Type = Double, Dynamic = False, Default = \"-2147483648", Scope = Private
	#tag EndConstant


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
End Module
#tag EndModule
