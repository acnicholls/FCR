<%@ Language=VBScript %>



<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->

<% 

dim  strSQL, strManu, strMake, strType, objCmd, intCarID
dim  intYear, strColour, blnAuto, intDoors, blnAC, intRate, intKM, strLocation
if request("cboLocation") > "" AND _
		request("txtmanu") > ""  AND _
		Request("txtcolour") > "" AND _
		request("txtMake") > "" AND _
		request("cboType") > "" AND _ 
		csng(request("txtRate")) > 0 AND _
		cint(request("cboDoors")) > 0 AND _
		cint(request("txtYear")) > 0 AND _
		clng(request("txtKM")) > 0 then

		set gobjrs = server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT max(fldCarID) FROM tblCar"
		call gobjrs.Open(strSQL, gobjConn, adOpenForwardOnly, adLockReadOnly, adCmdText)
				intCarID = gobjrs.Fields(0) + 1
		gobjrs.Close
		
				strLocation = request("cboLocation")
				strManu = request("txtmanu")
				strColour= Request("txtcolour")
				strMake= request("txtMake")
				strType=request("cboType")
				intRate=csng(request("txtRate"))
				if request("chkAuto") = "on" then
					blnAuto=1
				else
					blnAuto=0
				end if
				if request("chkAC") = "on" then
					blnAC=1
				else
					blnAC=0
				end if
				intDoors=cint(request("cboDoors"))
				intYear=cint(request("txtYear"))
				numKM=cint(request("txtKM"))
				strSQL = "INSERT INTO tblCar (fldCarID,fldManuf,fldMake,fldType,fldYear,fldColor,fldAuto,fldDoors,fldAC,fldRate,fldKM,fldLocation) VALUES('" & intCarID & "','" & strManu & "','" & strMake & "','" & strType & "'," & intYear & ",'" & strColour & "','" & blnAuto & "'," & intDoors & ",'" & blnAC & "'," & intRate & "," & numKM & ",'" & strLocation & "')"
				set objCmd = server.CreateObject("ADODB.COMMAND")
				objCmd.ActiveConnection = gobjConn
				objCmd.commandtype = adCmdText
				objCmd.commandtext = strSQL				
				objCmd.execute 
				
				Response.Redirect("refreshhome.asp")
				set objCmd = nothing
				set gobjrs = nothing
				
	else			

				Session("InsertFailure") = True
				Response.Redirect("newcar.asp")
				
end if
%>