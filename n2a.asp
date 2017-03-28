<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/KDB.asp" -->
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%

Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_KDB_STRING
Command1.CommandText = "INSERT INTO dbo.GalleryItems (GalleryID, Caption, Sequence, ImageURL)  VALUES (?,?,?,? ) "
Command1.Parameters.Append Command1.CreateParameter("GID", 3, 1, 4, MM_IIF(request("ID"), CInt(request("ID")), Command1__GID & ""))
Command1.Parameters.Append Command1.CreateParameter("C", 201, 1, 150, MM_IIF(request("gi_Caption"), request("gi_Caption"), Command1__C & ""))
Command1.Parameters.Append Command1.CreateParameter("S", 3, 1, 4, MM_IIF(request("gi_Sequence"), request("gi_Sequence"), Command1__S & ""))
Command1.Parameters.Append Command1.CreateParameter("URL", 201, 1, 150, MM_IIF(request("gi_ImageURL"), request("gi_ImageURL"), Command1__URL & ""))
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>

<%

response.Redirect("n2.asp?ID=" & request("ID"))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>
