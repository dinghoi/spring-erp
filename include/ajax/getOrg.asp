<%@LANGUAGE="VBSCRIPT"%>
<%Response.ContentType="text/html;charset=euc-kr"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim result : result = ""
Dim sql
Dim orgColumn, sqlWhere
Dim srchType, srchCompany, srchBonbu, srchSaupbu, srchTeam

srchType			= replaceXSS(request("srchType"))		'//�˻�����(1:ȸ��,2:����,3:�����,4:��,5:����ó)
srchCompany	= replaceXSS(request("srchCompany"))	'//ȸ���
srchBonbu		= replaceXSS(request("srchBonbu"))		'//���θ�
srchSaupbu		= replaceXSS(request("srchSaupbu"))	'//����θ�
srchTeam			= replaceXSS(request("srchTeam"))		'//����

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If srchType="" Or srchType="1" Then
	orgColumn = "org_company"
	sqlWhere = " AND org_level = 'ȸ��' "
ElseIf srchType="" Or srchType="2" Then
	orgColumn = "org_bonbu"
	sqlWhere = " AND org_level = '����' AND org_company='" & srchCompany & "' "
ElseIf srchType="" Or srchType="3" Then
	orgColumn = "org_saupbu"
	sqlWhere = " AND org_level = '�����' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' "
ElseIf srchType="" Or srchType="4" Then
	orgColumn = "org_team"
	sqlWhere = " AND org_level = '��' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' AND org_saupbu='" & srchSaupbu & "' "
ElseIf srchType="" Or srchType="5" Then
	orgColumn = "org_name"
	sqlWhere = " AND org_level = '����ó' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' AND org_saupbu='" & srchSaupbu & "' AND org_team='" & srchTeam & "' "
End If

'//get company list
sql="SELECT org_code, " & orgColumn & " FROM emp_org_mst WHERE isNull(org_end_date) " & sqlWhere & " ORDER BY org_code ASC"

Rs.Open Sql, Dbconn, 1

If Not(Rs.BOF Or Rs.EOF) Then
	Do Until Rs.EOF
		If Trim(result&"")<>"" Then result = result & ","
		result = result & "{""orgCode"":""" & Rs("org_code") & """"
		result = result & ",""orgCompany"":""" & Rs("org_company") & """}"
		Rs.movenext
	Loop
End If

Rs.close : Set Rs = Nothing
Dbconn.close : Set Dbconn = Nothing

'//return json data
If Trim(result&"")<>"" Then
	result = "{""result"":[" & result & "]}"
End If

Response.write result

%>