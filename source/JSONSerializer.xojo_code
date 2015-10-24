#tag Module
Protected Module JSONSerializer
	#tag Method, Flags = &h21
		Private Function extractString(jsonMB As MemoryBlock, ByRef pos As Integer, quoteChar As Byte) As String
		  Dim result() As String
		  
		  skipWhitespace jsonMB, pos
		  
		  pos = pos + 1
		  
		  while jsonMB.Byte(pos) <> quoteChar ' "
		    result.Append Chr(jsonMB.Byte(pos))
		    pos = pos + 1
		  wend
		  pos = pos + 1
		  
		  return Join(result, "")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function extractToken(jsonMB As MemoryBlock, ByRef pos As Integer) As String
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
		Function IsJSONObject(doc As String) As Boolean
		  Dim result As Boolean
		  Dim tmpStr As String
		  
		  tmpStr = Trim(doc)
		  
		  result = (Left(tmpStr, 1) = "{") and (Right(tmpStr, 1) = "}")
		  
		  return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParseJSON(json As String) As JSONItem
		  Dim pos As Integer
		  Dim jsonObj As JSONItem
		  
		  pos = 0
		  jsonObj = parseObject(json, pos)
		  
		  return jsonObj
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function parseName(jsonMB As MemoryBlock, ByRef pos As Integer) As String
		  Dim pairName As String
		  
		  skipWhitespace jsonMB, pos
		  
		  ' parse name (if not already given)
		  
		  if (jsonMB.Byte(pos) = 34) or (jsonMB.Byte(pos) = 39) then ' " OR '
		    pairName = extractString(jsonMB, pos, jsonMB.Byte(pos)) ' quoted name
		  else
		    pairName = extractToken(jsonMB, pos) ' non-quoted name
		  end if
		  
		  return pairName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function parseObject(jsonMB As MemoryBlock, ByRef pos As Integer) As JSONItem
		  Dim jsonResult As new JSONItem
		  Dim name As String
		  Dim value As Variant
		  
		  skipWhitespace jsonMB, pos
		  
		  if pos < jsonMB.Size then
		    
		    if jsonMB.Byte(pos) = 123 then ' {
		      
		      pos = pos + 1
		      skipWhitespace jsonMB, pos
		      
		      while jsonMB.Byte(pos) <> 125 ' }
		        
		        name = parseName(jsonMB, pos) ' parse name
		        
		        skipWhitespace jsonMB, pos
		        
		        pos = pos + 1 // skip past color
		        
		        skipWhitespace jsonMB, pos
		        
		        value = parseValue(jsonMB, pos)
		        
		        jsonResult.Value(name) = value ' parse value
		        
		        if  jsonMB.Byte(pos) = 44 then ' ,
		          pos = pos + 1
		        end if
		        
		        skipWhitespace jsonMB, pos
		        
		      wend
		      
		      pos = pos + 1
		      
		    end if
		    
		  end if
		  
		  return jsonResult
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function parseValue(jsonMB As MemoryBlock, ByRef pos As Integer) As Variant
		  Dim result As Variant
		  Dim tmpVal As String
		  Dim arrItem As JSONItem
		  
		  select case jsonMB.Byte(pos)
		    
		  case 34, 39 ' " OR ' = string
		    
		    result = extractString(jsonMB, pos, jsonMB.Byte(pos))
		    
		  case 91 ' [ = array
		    
		    arrItem = new JSONItem
		    
		    pos = pos + 1
		    
		    skipWhitespace jsonMB, pos
		    
		    while jsonMB.Byte(pos) <> 93 ' ]
		      
		      arrItem.Append parseValue(jsonMB, pos)
		      
		      skipWhitespace jsonMB, pos
		      
		      if  jsonMB.Byte(pos) = 44 then ' ,
		        pos = pos + 1
		        skipWhitespace jsonMB, pos
		      end if
		      
		    wend
		    pos = pos + 1 ' skip past ]
		    
		    result = arrItem
		    
		  case 123 ' { = object
		    
		    result = parseObject(jsonMB, pos)
		    
		  case else
		    
		    tmpVal = extractToken(jsonMB, pos)
		    
		    if tmpVal = "true" then
		      
		      result = true
		      
		    elseif tmpVal = "false" then
		      
		      result = false
		      
		    elseif tmpVal = "null" then
		      
		      result = nil
		      
		    else
		      
		      ' numeric value
		      
		      if InStr(1, tmpVal, ".") > 0 then
		        
		        ' double
		        
		        result = CDbl(tmpVal)
		        
		      else
		        
		        ' integer
		        
		        result = CLong(tmpVal)
		        
		      end if
		      
		    end if
		    
		  end select
		  
		  return result
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

	#tag Method, Flags = &h0
		Function StripJSONBrackets(s As String) As String
		  Dim tmpStr As String
		  
		  tmpStr = Trim(s)
		  
		  if Len(tmpStr) >= 2 then
		    
		    if Left(tmpStr, 1) = "{" then
		      tmpStr = Right(tmpStr, Len(tmpStr) - 1)
		    end if
		    
		    if Right(tmpStr, 1) = "}" then
		      tmpStr = Left(tmpStr, Len(tmpStr) - 1)
		    end if
		    
		  end if
		  
		  return tmpStr
		  
		End Function
	#tag EndMethod


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
