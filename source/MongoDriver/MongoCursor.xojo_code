#tag Class
Protected Class MongoCursor
	#tag Method, Flags = &h0
		Sub Constructor(initClient As MongoDriver.MongoClient, initCollection As MongoDriver.MongoCollection = nil)
		  mIndex = -1
		  
		  mClient = initClient
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  ' close the cursor on the server if needed
		  
		  if ID > 0 then
		    mClient.KillCursors ID 
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function getNext() As String
		  if (mIndex + 1) <= Document.Ubound then
		    mIndex = mIndex + 1
		    return Document(mIndex)
		  elseif ID <> 9 then
		    Redim Document(-1)
		    mClient.GetMore me
		    mIndex = -1
		    if Document.Ubound >= 0 then
		      mIndex = mIndex + 1
		      return Document(mIndex)
		    end if
		  end if
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function hasNext() As Boolean
		  Dim moreDocs As Boolean
		  
		  if (mIndex + 1) <= Document.Ubound Then
		    moreDocs = true
		  else
		    if ID <> 0 then
		      moreDocs = true
		    else
		      moreDocs = false
		    end if
		  end if
		  
		  return moreDocs
		End Function
	#tag EndMethod


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mCollection
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mCollection = value
			End Set
		#tag EndSetter
		Collection As MongoDriver.MongoCollection
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		Document() As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return (Document.Ubound + 1)
			  
			End Get
		#tag EndGetter
		DocumentCount As Integer
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		FullCollectionName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ID As Int64
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mClient As MongoDriver.MongoClient
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCollection As MongoDriver.MongoCollection
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberReturned As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberToReturn As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		ResponseFlags As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		StartingFrom As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="DocumentCount"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FullCollectionName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
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
			Name="NumberReturned"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NumberToReturn"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ResponseFlags"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="StartingFrom"
			Group="Behavior"
			Type="Integer"
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
