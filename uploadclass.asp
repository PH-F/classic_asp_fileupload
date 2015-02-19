<%

Class PseudoRequestDictionary
	Private oDic, bBinaryReadDone, bContainsFile, dtStartTime, oEmptyPseudoStringList
	Private iTotalFormBytes, sBoundary, bt13, bt34, btFileName

	Private Sub Class_Initialize
		Set oDic = Server.CreateObject("Scripting.Dictionary")
		bBinaryReadDone = False
		bContainsFile = False
		dtStartTime = Timer
		iTotalFormBytes = 0
		Set oEmptyPseudoStringList = new PseudoStringList
		bt13 = ChrB(13)
		bt34 = ChrB(34)
		btFileName = ChrB(102) & ChrB(105) & ChrB(108) & ChrB(101) & ChrB(110) & ChrB(97) & ChrB(109) & ChrB(101) & ChrB(61) & ChrB(34)
	End Sub

	Private Sub Class_Terminate
		RemoveAll
		Set oDic = Nothing
		Set oEmptyPseudoStringList = Nothing
	End Sub

	Private Function Byte2String(s)
		Dim i
		For i = 1 to LenB(s)
			Byte2String = Byte2String & CHR(AscB(MidB(s,i,1)))
		Next
	End Function

	Private Function RSBinaryToString(ByVal btBinary)
		Dim oRS, iLen
		Set oRS = Server.CreateObject("ADODB.Recordset")
		iLen = LenB(btBinary)

		If iLen > 0 Then
			oRS.Fields.Append "mBinary", 201, iLen
			oRS.Open
			oRS.AddNew
			oRS("mBinary").AppendChunk btBinary
			oRS.Update
			RSBinaryToString = oRS("mBinary").Value
			oRS.Close
		Else
			RSBinaryToString = ""
		End If
		Set oRS = Nothing
	End Function

	Private Sub AddFile(ByVal sKey, ByVal btValue)
		Dim oFileLoader, element, i

		sKey = LCase(sKey)
		If oDic.Exists(sKey) Then
			oDic.Item(sKey).AddRaw btValue
			If oDic.Item(sKey).FileSizeNumber(oDic.Item(sKey).Count) > 0 Then bContainsFile = True
		Else
			Set oFileLoader = new PseudoStringList
			Set oDic.Item(sKey) = oFileLoader
			oFileLoader.AddRaw btValue
			If oFileLoader.FileSize > 0 Then bContainsFile = True
			Set oFileLoader = Nothing
		End If
	End Sub

	Private Function TimeDiff(dtLast, dtFirst)
		Dim iReturn
		iReturn = dtLast - dtFirst
		If iReturn < 0 Then iReturn = dtLast - dtFirst + (60*60*24)
		TimeDiff = 1000 * iReturn
	End Function

	Public Sub ReadQuerystring(ByVal s)
		Dim aSplit, element, aSplitElement
		aSplit = Split(Cstr("" & s),"&")
		For each element in aSplit
			aSplitElement = Split(element,"=")
			Add aSplitElement(LBound(aSplitElement)), aSplitElement(UBound(aSplitElement))
		Next
	End Sub

	Public Sub ReadRequest()
		Dim sRequestContentType, aRequestContentType, sEncoding
		Dim aSplit, aSplitElement, element, sRequestMethod, sFormValues, aFormValues
		Dim i, j, iStartPos, iEndPos, btFormField, btBoundary
		Dim sKey, sValue, bFile, btBinaryRequest
		Dim dtStartTime, iFirstQuotePos, iSecondQuote

		sRequestContentType = Cstr(" " & Request.ServerVariables("HTTP_CONTENT_TYPE"))
		aRequestContentType = Split(sRequestContentType, ";")
		sEncoding = Trim(aRequestContentType(LBound(aRequestContentType)))
		iTotalFormBytes = Request.TotalBytes

		Select Case LCase(sEncoding)

			Case "multipart/form-data"
				btBinaryRequest = Request.BinaryRead(Request.TotalBytes)

				iStartPos = 1
				iEndPos = InstrB(1,btBinaryRequest,bt13)
				btBoundary = MidB(btBinaryRequest,1,iEndPos-1)
				sBoundary = Byte2String(btBoundary)
				dtStartTime = Timer

				iStartPos = iEndPos+2
				iEndPos = InstrB(iStartPos, btBinaryRequest, btBoundary)

				Do While iEndPos > 0
					bFile = False
					btFormField = MidB(btBinaryRequest,iStartPos,iEndPos-iStartPos-2)

					iFirstQuotePos = InstrB(btFormField, bt34)
					iSecondQuote = InstrB(iFirstQuotePos+1,btFormField, bt34)
					sKey = LCase(Byte2String(MidB(btFormField,iFirstQuotePos+1,iSecondQuote-iFirstQuotePos-1)))
					If InstrB(btFormField, btFileName) = iSecondQuote + 3 Then bFile = True

					If bFile Then
						AddFile sKey, btFormField
					Else
						sValue = Byte2String(RightB(btFormField, LenB(btFormField)-iSecondQuote-4))
						Add sKey, Server.URLEncode(sValue)
					End If

					iStartPos = iEndPos+2+LenB(btBoundary)
					iEndPos = InstrB(iStartPos, btBinaryRequest, btBoundary)
				Loop
				bBinaryReadDone = True

			Case "application/x-www-form-urlencoded"
				sRequestMethod = Request.ServerVariables("REQUEST_METHOD")
				If LCase(sRequestMethod) = "get" Then
					sFormValues = Request.Querystring
				ElseIf LCase(sRequestMethod) = "post" Then
					If Request.TotalBytes > 80000 Then
						' a very large form which can not be handled by normal request, the size is a bit arbitrary
						btBinaryRequest = Request.BinaryRead(Request.Totalbytes)
						bBinaryReadDone = True
						sFormValues = RSBinaryToString(btBinaryRequest)
					Else
						sFormValues = Request.Form
					End If
				End If
				aSplit = Split(sFormValues,"&")
				For each element in aSplit
					aSplitElement = Split(element,"=")
					Add aSplitElement(LBound(aSplitElement)), aSplitElement(UBound(aSplitElement))
				Next

			Case "text/plain"
				' missing "&" to seperate values, not urlencoded, usually used for sending mail directly
				'	not really relevant, but i will work it out some day

		End Select
	End Sub

	Public Sub Add(ByVal sKey, sValue)
		Dim oStringList
		sKey = LCase(sKey)

		If oDic.Exists(sKey) Then
			oDic.Item(sKey).Add sValue
		Else
			Set oDic.Item(sKey) = new PseudoStringList
			oDic.Item(sKey).Add sValue
		End If
	End Sub

	Public Sub ReplaceItem(ByVal sKey, sValue)
		sKey = LCase(sKey)
		Remove(sKey)
		Set oDic.Item(sKey) = new PseudoStringList
		oDic.Item(sKey).Add sValue
	End Sub

	Public Sub Remove(sKey)
		oDic(Lcase(skey)).Destroy
        oDic.Remove sKey
	End Sub

	Public Sub RemoveAll()
        Dim element
        For each element in oDic
			oDic(element).Destroy
		Next
        oDic.RemoveAll
	End Sub

	Public Property Get Exists(sKey)
        Exists = oDic.Exists(LCase(sKey))
	End Property

	Public Property Get Count
		Count = oDic.Count
	End Property

	Public Property Get BinaryReadDone
		BinaryReadDone = bBinaryReadDone
	End Property

	Public Property Get Keys
		Keys = oDic.Keys
	End Property

	Public Property Get Version
		Version = "1.4.01.04"
	End Property

	Public Property Get ContainsFile
		ContainsFile = bContainsFile
	End Property

	Public Property Get TotalFormBytes
		TotalFormBytes = iTotalFormBytes
	End Property

	Public Property Get Item(ByVal sKey)
		sKey = LCase(sKey)
		If oDic.Exists(sKey) Then
			Set Item = oDic.Item(sKey)
		Else
			Set Item = oEmptyPseudoStringList
		End If
	End Property

	Public Property Get ItemCount(ByVal sKey)
		sKey = LCase(sKey)
		ItemCount = 0
		If oDic.Exists(sKey) Then ItemCount = oDic.Item(sKey).Count
	End Property

	Public Default Function Value()
		Dim element, aElements, i, j, aSubelements
		ReDim aElements(oDic.Count - 1)
		i = 0
		For each element in oDic.Keys
			ReDim aSubelements(oDic(element).Count - 1)
			For j = 1 to oDic(element).Count
				aSubelements(j-1) = element & "=" & oDic(element).RawItem(j)
			Next
			aElements(i) = Join(aSubelements,"&")
			i = i + 1
		Next
		Value = Join(aElements,"&")
	End Function

	Public Property Get ExistenceTime
		ExistenceTime = Cstr("" & TimeDiff(Timer, dtStartTime)) & " ms"
	End Property

	Public Property Get Boundary
		Boundary = sBoundary
	End Property

End Class




Class PseudoStringList
	Private aList, bt34, bt13, btContentType

	Private Sub Class_Initialize
		aList = Array()
		bt34 = ChrB(34)
		bt13 = ChrB(13)
		btContentType = ChrB(67) & ChrB(111) & ChrB(110) & ChrB(116) & ChrB(101) & ChrB(110) & ChrB(116) & ChrB(45) & ChrB(84) & ChrB(121) & ChrB(112) & ChrB(101) & ChrB(58)
	End Sub

	Private Sub Class_Terminate
		Destroy
	End Sub

	Public Sub Destroy()
		Dim element
		For each element in aList
			If TypeName(element) = "Dictionary" Then
				element.RemoveAll
				Set element = Nothing
			End If
		Next
		Erase aList
	End Sub

	Public Sub Add(ByVal sString)
		ReDim Preserve aList(Ubound(aList)+1)
		aList(Ubound(aList)) = Cstr("" & sString)
	End Sub

	Public Property Get Count
		Count = UBound(aList) + 1
	End Property

	Public Property Get Keys
		Dim  aReturn, oDic, i
		If UBound(aList) => 0 Then
			ReDim aReturn(Ubound(aList))
			For i = 0 To Ubound(aList)
				If IsObject(aList(i)) Then
					aReturn(i) = aList(i)("filenamecomplete")
				Else
					aReturn(i) = URLDecode(aList(i))
				End If
			Next
			Keys = aReturn
		Else
			Keys = aList
		End If
	End Property

	Public Default Property Get Value
		Dim element, aElements, i
		Value = null
		If Ubound(aList) >= 0 Then
			Value = ""
			ReDim aElements(Ubound(aList))
			For i = Lbound(aList) To UBound(aList)
				If IsObject(aList(i)) Then
					aElements(i) = aList(i).Item("filenamecomplete")
				Else
					aElements(i) = URLDecode(aList(i))
				End If
			Next
			Value = Join(aElements,", ")
		End If
	End Property

	Public Property Get RawItem(i)
		If i > Ubound(aList)+1 Then Exit Property
		If IsObject(aList(i-1)) Then
			RawItem = Server.URLEncode(aList(i-1)("filenamecomplete"))
		Else
			RawItem = aList(i-1)
		End If
	End Property

	Public Property Get Version
		Version = "1.4.01.04"
	End Property

	Public Property Get ContainsFile
		Dim i
		ContainsFile = False
		For i = 0 To Ubound(aList)
			If TypeName(aList(i)) = "Dictionary" Then
				If aList(i)("filesize") > 0 Then
					ContainsFile = True
					Exit For
				End If
			End If
		Next
	End Property

	Public Property Get Item(i)
		If i > Ubound(aList)+1 Then Exit Property
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set Item = aList(i-1)
		Else
			Item = URLDecode(aList(i-1))
		End If
	End Property

	Private Sub AddFile(ByVal oDic)
		ReDim Preserve aList(Ubound(aList)+1)
		Set aList(Ubound(aList)) = oDic
	End Sub

	Public Sub AddRaw(ByVal btValue)
		Dim oDic, bFile, sFileName, sContentType
		Dim iFirstFoundPos, iSecondFoundPos

		Set oDic = Server.CreateObject("Scripting.Dictionary")
		iFirstFoundPos = InstrB(btValue, bt34)
		iSecondFoundPos = InstrB(iFirstFoundPos+1,btValue, bt34)

		iFirstFoundPos = InstrB(iSecondFoundPos+1,btValue, bt34)
		iSecondFoundPos = InstrB(iFirstFoundPos+1,btValue, bt34)

		If iSecondFoundPos > iFirstFoundPos + 1 Then bFile = True

		If bFile Then
			sFileName = Byte2String(MidB(btValue,iFirstFoundPos+1,iSecondFoundPos-iFirstFoundPos-1))
			oDic.Item("filenamecomplete") = sFileName
			' it is not sure if the file comes from Unix or Windows ("/" or "\"), don't know about Mac
			If Instr(sFileName,"\") > 0 Then		' suppose it is windows
				oDic.Item("filename") = Mid(sFileName, 1 + InStrRev(sFileName, "\"))
				oDic.Item("filepath") = Mid(sFileName, 1, InStrRev(sFileName, "\"))
			ElseIf Instr(sFileName,"/") > 0 Then
				oDic.Item("filename") = Mid(sFileName, 1 + InStrRev(sFileName, "/"))
				oDic.Item("filepath") = Mid(sFileName, 1, InStrRev(sFileName, "/"))
			Else
				' some browsers (Mozilla engine) do not have complete path
				oDic.Item("filename") = sFileName
				oDic.Item("filepath") = ""
			End If
on error resume next
			iFirstFoundPos = InstrB(iSecondFoundPos,btValue,btContentType)
			iSecondFoundPos = InstrB(iFirstFoundPos,btValue,bt13)
			sContentType = Byte2String(MidB(btValue,iFirstFoundPos+14,iSecondFoundPos-iFirstFoundPos-14))
			oDic("contenttype") = sContentType

			iFirstFoundPos = iSecondFoundPos+3
			oDic("binary") = RightB(btValue,LenB(btValue)-iFirstFoundPos)
			oDic("filesize") = LenB(oDic("binary"))

			AddFile(oDic)
		Else
			oDic.Item("filenamecomplete")= ""
			oDic.Item("filename")= ""
			oDic.Item("filepath")= ""
			oDic.Item("contenttype")= ""
			oDic("binary") = null
			oDic("filesize") = 0
			AddFile(oDic)
		End If

	End Sub

	Public Property Get FileName
		Dim oDic
		If TypeName(aList(0)) = "Dictionary" Then
			Set oDic = aList(0)
			FileName = oDic("filename")
			Set oDic = Nothing
		Else
			FileName = ""
		End If
	End Property

	Public Property Get FileSize
		Dim oDic
		If TypeName(aList(0)) = "Dictionary" Then
			Set oDic = aList(0)
			FileSize = oDic("filesize")
			Set oDic = Nothing
		Else
			FileSize = 0
		End If
	End Property

	Public Property Get IsFile(i)
		IsFile = False
		If i > Ubound(aList)+1 Then Exit Property
		If TypeName(aList(i-1)) = "Dictionary" Then IsFile = True
	End Property

	Public Property Get ContentType
		Dim oDic
		If TypeName(aList(0)) = "Dictionary" Then
			Set oDic = aList(0)
			ContentType = oDic("contenttype")
			Set oDic = Nothing
		Else
			ContentType = ""
		End If
	End Property

	Public Property Get Binary
		Dim oDic
		If TypeName(aList(0)) = "Dictionary" Then
			Set oDic = aList(0)
			Binary = oDic("binary")
			Set oDic = Nothing
		End If
	End Property

	Public Property Get BinaryNumber(i)
		Dim oDic
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set oDic = aList(i-1)
			Binary = oDic("binary")
			Set oDic = Nothing
		End If
	End Property

	Public Property Get ItemNumber(i)
		If i > Ubound(aList)+1 Then Exit Property
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set ItemNumber = aList(i-1)
		Else
			ItemNumber = aList(i-1)
		End If
	End Property

	Public Property Get ValueNumber(i)
		If i > Ubound(aList)+1 Then Exit Property
		If TypeName(aList(i-1)) = "Dictionary" Then
			ValueNumber = aList(i-1)("filenamecomplete")
		Else
			ValueNumber = URLDecode(aList(i-1))
		End If
	End Property

	Public Sub SaveAs(ByVal sFilePath)
		If Not IsNull(Binary) Then WriteBinaryFile sFilePath, Binary
	End Sub

	Private Function RSString2Byte(ByVal s)
		Dim iLenString, oRS
		Set oRS = Server.CreateObject("ADODB.Recordset")
		iLenString = LenB(s)
		If iLenString > 0 Then
			oRS.Fields.Append "mBinary", 205, iLenString
			oRS.Open
			oRS.AddNew
			oRS("mBinary").AppendChunk s & ChrB(0)
			oRS.Update
			RSString2Byte = oRS("mBinary").GetChunk(iLenString)
			oRS.Close
		End If
		Set oRS = Nothing
	End Function

	Private Sub WriteBinaryFile(sFilePath, ByVal sStream)
		Dim oStream
		Set oStream = Server.CreateObject("ADODB.Stream")
		oStream.Type = 1
		oStream.Open
		'response.write(sFilePath)
		oStream.Write RSString2Byte(sStream)

		On Error resume next
		
		oStream.SaveToFile sFilePath, 2

		if NOT Err.Number =0 then
		response.write("<H1>Fout, bestand bestaat al.</H1>")
		response.end
		end if

		on error goto 0

		oStream.Close
		Set oStream = Nothing
	End Sub

	Public Property Get FileNameNumber(i)
		Dim oDic
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set oDic = aList(i-1)
			FileNameNumber = oDic("filename")
			Set oDic = Nothing
		Else
			FileNameNumber = ""
		End If
	End Property

	Public Property Get FileSizeNumber(i)
		Dim oDic
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set oDic = aList(i-1)
			FileSizeNumber = oDic("filesize")
			Set oDic = Nothing
		Else
			FileSizeNumber = 0
		End If
	End Property

	Public Property Get ContentTypeNumber(i)
		Dim oDic
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set oDic = aList(i-1)
			ContentTypeNumber = oDic("contenttype")
			Set oDic = Nothing
		Else
			ContentTypeNumber = ""
		End If
	End Property

	Public Property Get ContainsFileNumber(i)
		ContainsFileNumber = False
		If Ubound(aList) >= 0 Then
			If TypeName(aList(i-1)) = "Dictionary" Then
				If aList(i-1)("filesize") > 0 Then ContainsFileNumber = True
			End If
		End If
	End Property

	Public Sub SaveAsNumber(sFilePath,i)
		Dim dTemp, bBinary
		If TypeName(aList(i-1)) = "Dictionary" Then
			Set dTemp = aList(i-1)
			bBinary = dTemp.Item("binary")
			WriteBinaryFile sFilePath, dTemp.Item("binary")
			Set dTemp = Nothing
		End If
	End Sub

	Private Function URLDecode(ByVal v)
		URLDecode = Replace(v,"+"," ")
		URLDecode = Unescape(URLDecode)
	End Function

	Private Function Byte2String(s)
		Dim i
		For i = 1 to LenB(s)
			Byte2String = Byte2String & CHR(AscB(MidB(s,i,1)))
		Next
	End Function
End Class

Function SaveFileAs(oPseudoRequest, sItem, sSaveDirectory, u)
	SaveFileAs = False
	'On Error Resume Next
	oPseudoRequest.Item(sItem).SaveAs(sSaveDirectory & u & oPseudoRequest.Item(sItem).FileName)
	If Not Err Then SaveFileAs = True
End Function


Function SaveFileAsWithNumber(oPseudoRequest, sItem, sSaveDirectory, i)
	SaveFileAsWithNumber = False
	'On Error Resume Next
	oPseudoRequest.Item(sItem).SaveAsNumber sSaveDirectory & oPseudoRequest.Item(sItem).FileNameNumber(i), i
	If Not Err Then SaveFileAsWithNumber = True
End Function

%>