<%
'list
Set objCmd = Server.CreateObject("ADODB.Command")
With objCmd
    .ActiveConnection = DBConn
    .CommandText = "Proc Name"
    .CommandType = adCmdStoredProc
End With
Set rsOrg = objCmd.Execute()
Set objCmd = Nothing

'return value

Set objCmd = Server.CreateObject("ADODB.Command")
With objCmd
    .ActiveConnection = DBConn
    .CommandText = "USP_ORG_SELECT_LEVEL"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@intResult",adInteger,adParamInput,4, 1)
    .Parameters.Append .CreateParameter("@strName",advarwchar,adParamInput,20, "����ȣ")
    .Parameters.Append .CreateParameter("@strTel",advarwchar,adParamInput,15, "010-4240-xxxx")
    .Parameters.Append .CreateParameter("@strEmail",advarwchar,adParamInput,50, "test�Դϴ�.")
    .Parameters.Append .CreateParameter("@strRegID",advarwchar,adParamInput,15, "test")

    .Parameters.Append .CreateParameter("@intResult",adInteger,adParamOutPut,0)

    '	.Parameters.Refresh
    '.Execute , , adExecuteNoRecords
End With
Set rsOrg = objCmd.Execute()
Set objCmd = Nothing
%>
