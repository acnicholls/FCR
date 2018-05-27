<% @Language=vbscript %>


<% Option Explicit %>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->

<%



dim intKM, intCarID, strLocation, strSQL, fileobj, strFileName, textfileObj, strClientID
set fileobj = server.CreateObject( "Scripting.FileSystemObject" )
strFileName = "C:\inetpub\wwwroot\fcr\App_Data\logs\" & "returnlog.txt"
intCarID = request("strCarID")
intKM = request("intKM")
strLocation = request("cboRetLocation")

strSQL = "SELECT * FROM tblCar where fldcarID = '" & intCarID & "'"
call gobjrs.open( strSQL, gobjConn, adopendynamic, adlockoptimistic)

gobjrs.fields(10) = null
strClientID = gobjrs.fields(13)
gobjrs.fields(13) = null
gobjrs.fields(12) = strLocation
gobjrs.fields(11) = intKM





if fileobj.FileExists( strFileName ) <> true then
	call fileobj.CreateTextFile( strFileName )
end if

	set textfileObj = fileobj.OpenTextFile( strFileName,8,true )
	
	call textfileobj.writeline("Car  " & intCarID)
	call textfileobj.writeline("Client  " & strClientID)
	call textfileobj.writeline("KM on Return   " & intKM)
	call textfileobj.writeline("Date Returned   " & date())
	call textfileobj.writeline("Returned to " & strLocation)
	call textfileobj.writeline("Operator  " & session("UserName"))
	call textfileobj.writeline("***********************************")
	
	call textfileobj.close
gobjrs.update
gobjrs.close
'now write to text file
Response.Redirect("refreshhome.asp")



%>

