<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/KDB.asp" -->
<%
Dim rsDetail__MMColParam
rsDetail__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsDetail__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsDetail
Dim rsDetail_cmd
Dim rsDetail_numRows

Set rsDetail_cmd = Server.CreateObject ("ADODB.Command")
rsDetail_cmd.ActiveConnection = MM_KDB_STRING
rsDetail_cmd.CommandText = "SELECT * FROM dbo.GalleryItems WHERE GalleryID = ?" 
rsDetail_cmd.Prepared = true
rsDetail_cmd.Parameters.Append rsDetail_cmd.CreateParameter("param1", 5, 1, -1, rsDetail__MMColParam) ' adDouble

Set rsDetail = rsDetail_cmd.Execute
rsDetail_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsDetail_numRows = rsDetail_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsDetail_total
Dim rsDetail_first
Dim rsDetail_last

' set the record count
rsDetail_total = rsDetail.RecordCount

' set the number of rows displayed on this page
If (rsDetail_numRows < 0) Then
  rsDetail_numRows = rsDetail_total
Elseif (rsDetail_numRows = 0) Then
  rsDetail_numRows = 1
End If

' set the first and last displayed record
rsDetail_first = 1
rsDetail_last  = rsDetail_first + rsDetail_numRows - 1

' if we have the correct record count, check the other stats
If (rsDetail_total <> -1) Then
  If (rsDetail_first > rsDetail_total) Then
    rsDetail_first = rsDetail_total
  End If
  If (rsDetail_last > rsDetail_total) Then
    rsDetail_last = rsDetail_total
  End If
  If (rsDetail_numRows > rsDetail_total) Then
    rsDetail_numRows = rsDetail_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsDetail
MM_rsCount   = rsDetail_total
MM_size      = rsDetail_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsDetail_first = MM_offset + 1
rsDetail_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsDetail_first > MM_rsCount) Then
    rsDetail_first = MM_rsCount
  End If
  If (rsDetail_last > MM_rsCount) Then
    rsDetail_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>


<link href="bootstrap/css/bootstrap.css" rel="stylesheet" type="text/css" />



<script language="javascript">

</script>
</head>

<body>


<div class="container">

  <div class="well">Gallery Detail </div>
  <table class="table table-striped">
    <tr>
      <td>GalleryItemID</td>
      <td>GalleryID</td>
      <td>Caption</td>
      <td>Sequence</td>
      <td>ImageURL</td>
      <td>Action</td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsDetail.EOF)) %>
      <tr>
        <td><%=(rsDetail.Fields.Item("GalleryItemID").Value)%></td>
        <td><%=(rsDetail.Fields.Item("GalleryID").Value)%></td>
        <td><%=(rsDetail.Fields.Item("Caption").Value)%></td>
        <td><%=(rsDetail.Fields.Item("Sequence").Value)%></td>
        <td><%=(rsDetail.Fields.Item("ImageURL").Value)%></td>
        <td><a href="dgi.asp?GID=<%=(rsDetail.Fields.Item("GalleryID").Value)%>&ID=<%=(rsDetail.Fields.Item("GalleryItemID").Value)%>" class="btn btn-primary">Delete</a></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsDetail.MoveNext()
Wend
%>
  </table>
  <table border="0">
    <tr>
      <td><% If MM_offset <> 0 Then %>
          <a href="<%=MM_moveFirst%>">First</a>
          <% End If ' end MM_offset <> 0 %></td>
      <td><% If MM_offset <> 0 Then %>
          <a href="<%=MM_movePrev%>">Previous</a>
          <% End If ' end MM_offset <> 0 %></td>
      <td><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveNext%>">Next</a>
          <% End If ' end Not MM_atTotal %></td>
      <td><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveLast%>">Last</a>
          <% End If ' end Not MM_atTotal %></td>
    </tr>
  </table>
  
  
  <div class="well">New Gallery Item</div>
  
  
<form action="n2a.asp" method="post">
<div class="form-group">
<label>Caption</label>
<input type="text" class="form-control" name="gi_Caption" id="gi_Caption" value="" placeholder="Enter caption"  />
</div>
<div class="form-group">
<label>Sequence</label>
<input type="text" class="form-control" name="gi_Sequence" id="gi_Sequence" value="" placeholder="Order sequence"  />
</div>

<div class="form-group">
<label>ImageURL</label>
<input type="text" class="form-control" name="gi_ImageURL" id="gi_ImageURL" value="" placeholder="Enter URL"  />
<input type="hidden" id="ID" name="ID" value="<%=request("ID")%>" />

</div>





<button type="submit" class="btn btn-default">Submit</button>

</form>

<p>

<div class="btn-group">
  <a href="n1.asp" class="btn btn-primary">All Galleries</a>
</div>


<div class="well">Gallery Browse</div>
<%

' create gallery of all items in folder

Set fs=Server.CreateObject("Scripting.FileSystemObject")
set folder =  fs.GetFolder(Server.MapPath("Gallery"))
set files  = folder.Files
if files.Count > 0 then 
dim fc 

fc = 1

for each file in files

%>



<% if fc = 1 then %>
<div class="row">
 
  
<% end if %>
 <div class="col-md-4">
		<div class="thumbnail">
			<img src="Gallery/<%=file.name%>" class="img-thumbnail" />
			<div class="caption">
    			<p><%=fc%>&nbsp;<%=file.name%></p>
	    	</div>
       	</div>
 </div>      

<%  fc = fc + 1 
if fc = 4 then 
 fc = 1
end if

next 

end if %>

</div>
</body>
</html>
<%
rsDetail.Close()
Set rsDetail = Nothing
%>
