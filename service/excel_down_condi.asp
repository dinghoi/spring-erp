<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">-->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim from_date, to_date, company, date_sw, process_sw
Dim rs_etc, rs

Dim field_check, field_view, savefilename
Dim k, i

'Dim whatever
'Dim alldata
'Dim company_tab(50)

'title_name = array("������ȣ", "��������", "������", "�����", "��������", "��ȭ��ȣ", "�ڵ���", "ȸ��", "������", "�ּ�", "CE��", "CE���", "CE�Ҽ���", "��ֳ���", "��û��", "��û�ð�", "ó����", "ó���ð�", "����", "ó�����", "����û", "�԰�/��������", "�԰�����", "��ü����", "����Ŀ", "������", "�ڻ��ڵ�", "�𵨸�", "S/N��ȣ", "ó������", "��ġ����", "PC S/W", "PC H/W", "�����", "������/���ɳ�", "������", "����/��ũ", "�ƴ���", "��Ÿ")

from_date = Request("from_date")
to_date = Request("to_date")
company = Request("company")
date_sw = Request("date_sw")
process_sw = Request("process_sw")
field_check = Request("field_check")
field_view = Request("field_view")
savefilename = from_date & to_date & ".xls"

Response.Buffer = True
Response.Expires = 0
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
'Response.AddHeader "Content-Type", "application/json; charset=utf-8"
Response.AddHeader "Content-Disposition","attachment; filename=" & savefilename

If c_grade = "7" Then
	k = 0

	'Sql="SELECT * FROM etc_code WHERE etc_type = '51' AND used_sw = 'Y' AND group_name = '"+user_name+"' ORDER BY etc_name ASC"
	objBuilder.Append "SELECT etc_name FROM etc_code "
	objBuilder.Append "WHERE etc_type = '51' AND used_sw = 'Y' AND group_name = '"&user_name&"' "
	objBuilder.Append "ORDER BY etc_name ASC "

	rs_etc.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	While Not rs_etc.EOF
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.MoveNext()
	Wend
	rs_etc.close()
End If

' 2018-03-06 as_acpt.mg_ce_id �Ҽ� ���� ǥ�� from emp_master
objBuilder.Append "SELECT asat.acpt_no, asat.acpt_date, asat.acpt_man, asat.acpt_user, "
objBuilder.Append "	CASE WHEN IFNULL(asat.cowork_yn, 'N') = 'N' THEN 'NO' WHEN IFNULL(asat.cowork_yn, 'N') = 'Y' THEN 'YES' END AS cowork, "
objBuilder.Append "	CONCAT(asat.tel_ddd, '-', asat.tel_no1, '-', asat.tel_no2) AS as_tel, "
objBuilder.Append "	CONCAT(asat.hp_ddd, '-', asat.hp_no1, '-', asat.hp_no2) AS as_hp, asat.company, asat.dept, "
objBuilder.Append "	CONCAT(asat.sido,' ', asat.gugun, ' ', asat.dong, ' ', asat.addr) AS as_address, asat.mg_ce, asat.mg_ce_id, "
objBuilder.Append "	asat.as_memo, asat.request_date, asat.request_time, asat.visit_date, asat.visit_time, asat.as_process, "
objBuilder.Append "	asat.as_type, asat.visit_request_yn, asat.into_reason, asat.in_date, asat.in_replace, asat.maker, "
objBuilder.Append "	asat.as_device, asat.asets_no, asat.model_no, asat.serial_no, asat.as_history, asat.dev_inst_cnt, "
objBuilder.Append "	asat.err_pc_sw, asat.err_pc_hw, asat.err_monitor, asat.err_printer, asat.err_network, "
objBuilder.Append "	asat.err_server, asat.err_adapter, asat.err_etc, "
objBuilder.Append "emtt.emp_org_name "
objBuilder.Append "FROM as_acpt AS asat "
objBuilder.Append "INNER JOIN emp_master AS emtt ON asat.mg_ce_id = emtt.emp_no "

If date_sw = "acpt" Then
	'date_sql = "where (cast(A.acpt_date as date) >= '" + from_date  + "' and cast(A.acpt_date as date) <= '" + to_date  + "')"
	objBuilder.Append "WHERE (CAST(asat.acpt_date AS DATE) >= '"&from_date&"' AND CAST(asat.acpt_date AS DATE) <= '"&to_date&"') "
Else
	'date_sql = "where (A.visit_date >= '" + from_date  + "' and A.visit_date <= '" + to_date  + "')"
	objBuilder.Append "WHERE (asat.visit_date >= '"&from_date&"' AND asat.visit_date <= '"&to_date&"') "
End If

If c_grade = "7" Then
	com_sql = "asat.company = '"&company_tab(1)&"' "
	For kk = 2 To k
		com_sql = com_sql & " OR asat.company = '"&company_tab(kk)&"'"
	Next

	'sql = base_sql + date_sql + " and (" + com_sql + ") " + process_sql + field_sql
	objBuilder.Append "AND ("&com_sql&") "
End If

If c_grade = "8" Then
	com_sql = "AND (company = '"&user_name&"') "
	'sql = base_sql + date_sql + com_sql + process_sql + field_sql
	objBuilder.Append com_sql
End If

If company <> "��ü" Then
	objBuilder.Append "AND (asat.company = '"&company&"') "
End If

If process_sw = "A" Then
	'process_sql = " and ( A.as_process = '�Ϸ�' or A.as_process = '��ü' or A.as_process = '���' or A.as_process = '����' or A.as_process = '����' or A.as_process = '�԰�' or as_process = '��ü�԰�' ) "
	objBuilder.Append "AND (asat.as_process = '�Ϸ�' OR asat.as_process = '��ü' OR asat.as_process = '���' "
	objBuilder.Append "OR asat.as_process = '����' OR asat.as_process = '����' OR asat.as_process = '�԰�' OR asat.as_process = '��ü�԰�') "
ElseIf process_sw = "Y" Then
	'process_sql = " and ( A.as_process = '�Ϸ�' or A.as_process = '��ü' or A.as_process = '���') "
	objBuilder.Append "AND (asat.as_process = '�Ϸ�' OR asat.as_process = '��ü' OR asat.as_process = '���') "
Else
	'process_sql = " and ( A.as_process = '����' or A.as_process = '����' or A.as_process = '�԰�' or as_process = '��ü�԰�') "
	objBuilder.Append "AND (asat.as_process = '����' OR asat.as_process = '����' OR asat.as_process = '�԰�' OR asat.as_process = '��ü�԰�')"
End If

If field_check <> "total" Then
	'field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) ORDER BY A.acpt_date DESC"
	objBuilder.Append "AND ("&field_check&" LIKE '%"&field_view&"%') ORDER BY asat.acpt_date DESC"
Else
	'field_sql = " ORDER BY A.acpt_date DESC"
	objBuilder.Append "ORDER BY asat.acpt_date DESC "
End If

'If company = "��ü" Then
'	sql = sql
'Else
'	com_sql = " and (A.company = '" + company + "') "
'	sql = base_sql + date_sql + com_sql + process_sql + field_sql
'End If

'sql = base_sql + date_sql + process_sql + field_sql

'Response.write objBuilder.ToString()
'Response.end

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

If rs.EOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('�ٿ� �� �ڷᰡ �����ϴ�.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
End If

Dim rsErrSw, rsErrHw, rsErrMonitor, rsErrPrint, rsErrNw
Dim rsErrServer, rsErrAdapt, rsErrEtc

'Set rsErrSw = Server.CreateObject("ADODB.RecordSet")
'Set rsErrHw = Server.CreateObject("ADODB.RecordSet")
'Set rsErrMonitor = Server.CreateObject("ADODB.RecordSet")
'Set rsErrPrint = Server.CreateObject("ADODB.RecordSet")
'Set rsErrNw = Server.CreateObject("ADODB.RecordSet")
'Set rsErrServer = Server.CreateObject("ADODB.RecordSet")
'Set rsErrAdapt = Server.CreateObject("ADODB.RecordSet")
'Set rsErrEtc = Server.CreateObject("ADODB.RecordSet")

%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<title></title>
	</head>
	<body>
		<!--<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>-->
		<table border='1' cellspacing='0' cellpadding='5'>
			<tr>
				<td>������ȣ</td>
				<td>��������</td>
				<td>������</td>
				<td>�����</td>
				<td>��������</td>
				<td>��ȭ��ȣ</td>
				<td>�޴���</td>
				<td>ȸ��</td>
				<td>������</td>
				<td>�ּ�</td>
				<td>CE��</td>
				<td>CE���</td>
				<td>CE�Ҽ���</td>
				<td>��ֳ���</td>
				<td>��û��</td>
				<td>��û�ð�</td>
				<td>ó����</td>
				<td>ó���ð�</td>
				<td>����</td>
				<td>ó�����</td>
				<td>����û</td>
				<td>�԰�/��������</td>
				<td>�԰�����</td>
				<td>��ü����</td>
				<td>����Ŀ</td>
				<td>������</td>
				<td>�ڻ��ڵ�</td>
				<td>�𵨸�</td>
				<td>S/N��ȣ</td>
				<td>ó������</td>
				<td>��ġ����</td>
				<td>PC S/W</td>
				<td>PC H/W</td>
				<td>�����</td>
				<td>������/���ɳ�</td>
				<td>������</td>
				<td>����/��ũ</td>
				<td>�ƴ���</td>
			</tr>
			<%
			Do Until rs.EOF
			%>
			<tr>
				<td><%=rs("acpt_no")%></td>
				<td><%=rs("acpt_date")%></td>
				<td><%=rs("acpt_man")%></td>
				<td><%=rs("acpt_user")%></td>
				<td><%=rs("cowork")%></td>
				<td><%=rs("as_tel")%></td>
				<td><%=rs("as_hp")%></td>
				<td><%=rs("company")%></td>
				<td><%=rs("dept")%></td>
				<td><%=rs("as_address")%></td>
				<td><%=rs("mg_ce")%></td>
				<td><%=rs("mg_ce_id")%></td>
				<td><%=rs("emp_org_name")%></td>
				<td><%=rs("as_memo")%></td>
				<td><%=rs("request_date")%></td>
				<td><%=rs("request_time")%></td>
				<td><%=rs("visit_date")%></td>
				<td><%=rs("visit_time")%></td>
				<td><%=rs("as_process")%></td>
				<td><%=rs("as_type")%></td>
				<td>
				<%	'����û
				If rs("visit_request_yn") = "Y" Then
					Response.Write "�湮��û"
				Else
					Response.Write ""
				End If
				%>
				</td>
				<td><%=rs("into_reason")%></td>
				<td><%=rs("in_date")%></td>
				<td><%=rs("in_replace")%></td>
				<td><%=rs("maker")%></td>
				<td><%=rs("as_device")%></td>
				<td><%=rs("asets_no")%></td>
				<td><%=rs("model_no")%></td>
				<td><%=rs("serial_no")%></td>
				<td><%=rs("as_history")%></td>
				<td><%=rs("dev_inst_cnt")%></td>
				<td>
				<%
				If IsNull(rs("err_pc_sw")) Or rs("err_pc_sw") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_pc_sw"))&") "

					'rsErrSw.Open objBuilder.ToString(), DBConn, 1
					Set rsErrSw = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrSw.EOF
					'	Response.Write "- " & rsErrSw("etc_name") & "<br/>"
					'	rsErrSw.MoveNext()
					'Loop
					Response.Write rsErrSw("err_name")

					rsErrSw.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_pc_hw")) Or rs("err_pc_hw") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_pc_hw"))&") "

					'rsErrHw.Open objBuilder.ToString(), DBConn, 1
					Set rsErrHw = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrHw.EOF
					'	Response.Write "- " & rsErrHw("etc_name") & "<br/>"
					'	rsErrHw.MoveNext()
					'Loop
					Response.Write rsErrHw("err_name")
					rsErrHw.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_monitor")) Or rs("err_monitor") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_monitor"))&") "

					'rsErrMonitor.Open objBuilder.ToString(), DBConn, 1
					Set rsErrMonitor = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrMonitor.EOF
					'	Response.Write "- " & rsErrMonitor("etc_name") & "<br/>"
					'	rsErrMonitor.MoveNext()
					'Loop
					Response.Write rsErrMonitor("err_name")
					rsErrMonitor.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_printer")) Or rs("err_printer") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_printer"))&") "

					'rsErrPrint.Open objBuilder.ToString(), DBConn, 1
					Set rsErrPrint = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrPrint.EOF
					'	Response.Write "- " & rsErrPrint("etc_name") & "<br/>"
					'	rsErrPrint.MoveNext()
					'Loop
					Response.Write rsErrPrint("err_name")
					rsErrPrint.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_network")) Or rs("err_network") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_network"))&") "

					'rsErrNw.Open objBuilder.ToString(), DBConn, 1
					Set rsErrNw = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrNw.EOF
					'	Response.Write "- " & rsErrNw("etc_name") & "<br/>"
					'	rsErrNw.MoveNext()
					'Loop
					Response.Write rsErrNw("err_name")
					rsErrNw.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_server")) Or rs("err_server") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_server"))&") "

					'rsErrServer.Open objBuilder.ToString(), DBConn, 1
					Set rsErrServer = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrServer.EOF
					'	Response.Write "- " & rsErrServer("etc_name") & "<br/>"
					'	rsErrServer.MoveNext()
					'Loop
					Response.Write rsErrServer("err_name")
					rsErrServer.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_adapter")) Or rs("err_adapter") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_adapter"))&") "

					'rsErrAdapt.Open objBuilder.ToString(), DBConn, 1
					Set rsErrAdapt = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrAdapt.EOF
					'	Response.Write "- " & rsErrAdapt("etc_name") & "<br/>"
					'	rsErrAdapt.MoveNext()
					'Loop
					Response.Write rsErrAdapt("err_name")
					rsErrAdapt.Close()
				End If
				%>
				</td>
				<td>
				<%
				If IsNull(rs("err_etc")) Or rs("err_etc") = "" Then
					Response.write ""
				Else
					objBuilder.Append "SELECT GROUP_CONCAT(etc_name SEPARATOR ',') AS err_name FROM etc_code "
					objBuilder.Append "WHERE etc_code IN ("&SetAsListExcelErrName(rs("err_etc"))&") "

					'rsErrEtc.Open objBuilder.ToString(), DBConn, 1
					Set rsErrEtc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					'Do Until rsErrEtc.EOF
					'	Response.Write "- " & rsErrEtc("etc_name") & "<br/>"
					'	rsErrEtc.MoveNext()
					'Loop
					Response.Write rsErrEtc("err_name")
					rsErrEtc.Close()
				End If
				%>
				</td>
			</tr>
			<%
				rs.MoveNext()
			Loop

			Set rs_etc = Nothing
			Set rsErrSw = Nothing
			Set rsErrHw = Nothing
			Set rsErrMonitor = Nothing
			Set rsErrPrint = Nothing
			Set rsErrNw = Nothing
			Set rsErrServer = Nothing
			Set rsErrAdapt = Nothing
			Set rsErrEtc = Nothing
			rs.Close() : Set rs = Nothing
			%>
			<!--<td style="mso-number-format:'\@'" valign="top"></td>-->
		<!--</tr><%'=Chr(13)&Chr(10)%>-->
		</table>
	</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->
