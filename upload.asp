<!--#include file="uploadclass.asp"-->
<%
Response.Buffer = False
Server.ScriptTimeOut = 300

Dim oPseudoRequest, isSaved, action

Set oPseudoRequest = new PseudoRequestDictionary
oPseudoRequest.ReadRequest()
oPseudoRequest.ReadQuerystring(Request.Querystring)

	action = oPseudoRequest.Item("action")
	If oPseudoRequest.Item("file_1").ContainsFile Then
		%>
		
		<table border=0>
		<tr>
			<td>file</td>
			<td><%=oPseudoRequest.Item("file_1")%></td>
		</tr>
		<tr>
			<td>name</td>
			<td><%=oPseudoRequest.Item("file_1").FileName%></td>
		</tr>
		<tr>
			<td>size</td>
			<td><%=oPseudoRequest.Item("file_1").FileSize%> bytes</TD>
		</TR>
		<tr>
			<td>type</td>
			<td><%=oPseudoRequest.Item("file_1").ContentType%></td>
		</tr>
		</table>
		<%
		path="c:\inetpub\site\upload\"

		isSaved = SaveFileAs(oPseudoRequest, "file_1", path, "")
		If isSaved Then
			If action = "test" Then
				response.write "opgeslagen"
			Else
				response.redirect "/"
			End If
		Else
			Response.write "Error saving file."
		End If
	Else
		Response.write "Please choose a file."
	End If
	
Set oPseudoRequest = Nothing
%>