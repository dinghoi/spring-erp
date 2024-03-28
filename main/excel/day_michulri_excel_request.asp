<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim com_tab(15)
Dim com_sum(15)
Dim ok_sum(15)
Dim mi_sum(15)
Dim com_cnt(15,7)
Dim sum_cnt(7)
Dim company_tab(150)
Dim end_tab(8)
Dim mi_tab(8)

Dim from_date, to_date, sido, mg_ce, company, as_type, days
Dim curr_day, curr_date, win_sw, mg_ce_id, memo01, memo02
Dim com_sql, type_sql, title_memo, savefilename
Dim rsAs, arrAs

from_date = f_Request("from_date")
to_date = f_Request("to_date")

sido = f_Request("sido")
mg_ce = f_Request("mg_ce")
mg_ce_id = f_Request("mg_ce_id")
mg_group = f_Request("mg_group")
company = f_Request("company")
as_type = f_Request("as_type")
days = Int(f_Request("days"))

curr_day = DateValue(Mid(CStr(Now()), 1, 10))
curr_date = DateValue(Mid(DateAdd("h", 12, Now()), 1, 10))

win_sw = "back"

If company = "" Then
	company = "��ü"
	as_type = "��ü"
End If

If mg_ce = "" Then
	memo01 = "�õ�"
	memo02 = sido
Else
	memo01 = "�����"
	memo02 = mg_ce
End If

If company = "��ü" Then
	com_sql = ""
Else
  	com_sql = "company ='"&company&"' AND "
End If

If as_type = "��ü" Then
	type_sql = ""
Else
  	type_sql = "as_type ='"&as_type&"' AND "
End If

If mg_ce = "" Then
	title_memo = sido & "_������_"
Else
    title_memo = mg_ce & "_�����_"
End If

savefilename = title_memo & CStr(days) & "�� ��û���� ���� ��ó�� ����.xls"

'Response.Buffer = True
'Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
'Response.CacheControl = "public"
'Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

'Call ViewExcelType(savefilename)

objBuilder.Append "SELECT acpt_no, request_date, as_process, company, dept, sido, gugun, as_type, acpt_man, "
objBuilder.Append "	tel_ddd, tel_no1, tel_no2, addr, mg_ce, as_memo, request_time, into_reason, "
objBuilder.Append "	acpt_user, dong, acpt_date "
objBuilder.Append "FROM as_acpt "
objBuilder.Append "WHERE "&com_sql&type_sql&" (as_process = '����' OR as_process = '�԰�' OR as_process = '����') "
objBuilder.Append "	AND (Cast(request_date as date) >= '"&from_date&"' AND Cast(request_date as date) <= '"&to_date&"') "

' ��ó����
If mg_ce = "" Then
	Select Case sido
		Case "�Ѱ�", "��"
			objBuilder.Append ""
		Case "����"
			objBuilder.Append " AND sido IN ('����', '���', '��õ') "
		Case "�λ�����"
			objBuilder.Append "	AND sido IN ('�λ�', '�泲', '���') "
		Case "�뱸����"
			objBuilder.Append "	AND sido IN ('�뱸', '���') "
		Case "��������"
			objBuilder.Append "	AND sido IN ('����', '�泲', '���', '����') "
			objBuilder.Append "	AND (GUGUN <> '��õ��' AND GUGUN <> '�ܾ籺') "	 ' �����õ�ÿ� �ܾ籺�� �������翡�� ��������� ������ ����� (2019.01.18)  ����� ���� �䱸
		Case "��������"
			objBuilder.Append "	AND sido IN ('����', '����', '����') "
		Case "��������"
			objBuilder.Append "	AND sido = '����' "
		Case "��������"
			objBuilder.Append "	AND sido = '����' "
			objBuilder.Append "	OR (GUGUN = '��õ��' OR GUGUN = '�ܾ籺') "	 ' �����õ�ÿ� �ܾ籺�� �������翡�� ��������� ������ ����� (2019.01.18)  ����� ���� �䱸
		Case Else
			objBuilder.Append "	AND sido = '"&sido&"' "
	End Select
Else
	If mg_ce <> "�Ѱ�" Then
		objBuilder.Append " AND mg_ce_id = '"&mg_ce_id&"' "
	End If
End If

Set rsAs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAs.EOF Then
	arrAs = rsAs.getRows()
End If
rsAs.Close() : Set rsAs = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title></title>
	<style type="text/css">
	<!--
	.style14 {color: #FFCCFF}
	-->
	</style>
</head>
<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0">
	<tr bgcolor="#CCCCCC" class="style11">
		<td height="25" bgcolor="#FFFFFF"><%=memo01%></td>
		<td height="25" bgcolor="#FFFFFF">&nbsp;<%=memo02%></td>
		<td height="25" bgcolor="#FFFFFF">ȸ��</td>
		<td height="25" bgcolor="#FFFFFF">&nbsp;<%=company%></td>
		<td height="25" bgcolor="#FFFFFF">ó������</td>
		<td height="25" bgcolor="#FFFFFF">&nbsp;<%=as_type%></td>
		<td height="25" bgcolor="#FFFFFF">�Ⱓ</td>
		<td bgcolor="#FFFFFF"><%=days%>�� ��ó��</td>
		<td bgcolor="#FFFFFF">&nbsp;</td>
		<td bgcolor="#FFFFFF">�������� ����</td>
		<td bgcolor="#FFFFFF">&nbsp;</td>
		<td bgcolor="#FFFFFF">&nbsp;</td>
		<td bgcolor="#FFFFFF">&nbsp;</td>
		<td bgcolor="#FFFFFF">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="style11">
		<td width="88"><div align="center"><strong>��������</strong></div></td>
		<td width="57" height="20"><div align="center"><strong><span class="style25">������</span></strong></div></td>
		<td width="56" height="20"><div align="center"><strong>�����</strong></div></td>
		<td width="101" height="20" class="style11B"><div align="center"><strong>��ȭ��ȣ</strong></div></td>
		<td width="102" height="20" class="style11B"><div align="center"><strong>ȸ��</strong></div></td>
		<td width="101" height="20" class="style11B"><div align="center"><strong>������</strong></div></td>
		<td width="165" height="20"><div align="center"><strong>�ּ�</strong></div></td>
		<td width="63"><div align="center"><strong>CE��</strong></div></td>
		<td width="110"><div align="center"><strong>��ֳ���</strong></div></td>
		<td width="64"><div align="center"><strong>��û��</strong></div></td>
		<td width="57"><div align="center"><strong>��û�ð�</strong></div></td>
		<td width="56"><div align="center"><strong>ó�����</strong></div></td>
		<td width="38"><div align="center"><strong>����</strong></div></td>
		<td width="55"><div align="center"><strong>�԰����</strong></div></td>
	</tr>
	<%
	Dim seq, as_acpt_no, as_request_date, as_process, as_company, as_dept
	Dim as_sido, as_gugun, as_as_type, as_acpt_man, as_tel_ddd, as_tel_no1, as_tel_no2
	Dim as_addr, as_mg_ce, as_memo, as_request_time, into_reason
	Dim l, com_date, dd, a, d, rs_week, rs_hol, acpt_day, ddd, curr_hh, acpt_hh
	Dim as_acpt_date, as_acpt_user, as_dong

	If IsArray(arrAs) Then
		seq = 0

		For l = LBound(arrAs) To UBound(arrAs, 2)
			as_acpt_no = arrAs(0, l)
			as_request_date = arrAs(1, l)
			as_process = arrAs(2, l)
			as_company = arrAs(3, l)
			as_dept = arrAs(4, l)
			as_sido = arrAs(5, l)
			as_gugun = arrAs(6, l)
			as_as_type = arrAs(7, l)
			as_acpt_man = arrAs(8, l)
			as_tel_ddd = arrAs(9, l)
			as_tel_no1 = arrAs(10, l)
			as_tel_no2 = arrAs(11, l)
			as_addr = arrAs(12, l)
			as_mg_ce = arrAs(13, l)
			as_memo = arrAs(14, l)
			as_request_time = arrAs(15, l)
			into_reason = arrAs(16, l)
			as_acpt_user = arrAs(17, l)
			as_dong = arrAs(18, l)
			as_acpt_date = arrAs(19, l)

			seq = seq + 1

			com_date = DateValue(Mid(DateAdd("h", 10, as_request_date), 1, 10))
			dd = DateDiff("d", com_date, curr_date)
			'				ddd = dd
			'���� ���
			If dd < 0 Then
				dd = 0
			End If

			If CStr(curr_day) = CStr(Mid(as_request_date, 1, 10)) Then
				dd = 0
			End If

			If dd > 0 Then
				com_date = DateValue(Mid(as_request_date, 1, 10))
				'a = dd
				a = DateDiff("d", com_date, curr_day)
				'b = datepart("w", com_date)
				'bb = datepart("w", curr_date)
				'if bb = 1 then
				'	a = a -1
				'end if
				'c = a + b
				d = a
				'if a > 1 then
				'	if c > 7 then
				'		d = a - 2
				'	end if
				'end if

				Do Until com_date > curr_day
					objBuilder.Append "SELECT dayweeks FROM (SELECT DAYOFWEEK('"&CStr(com_date)&"') AS dayweeks) A WHERE A.dayweeks IN (1,7) "
					Set rs_week = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					If rs_week.EOF Or rs_week.BOF Then
						d = d
					Else
						d = d - 1
					End If

					com_date = DateAdd("d", 1, com_date)
					rs_week.Close()
				Loop
				Set rs_week = Nothing

		'		visit_date = rs("visit_date")
		'					com_date = datevalue(rs("acpt_date"))
		'		act_date = com_date

				com_date = DateValue(Mid(as_request_date, 1, 10))

				Do Until com_date > curr_day
					objBuilder.Append "SELECT holiday FROM holiday WHERE holiday = '"&CStr(com_date)&"' "
					Set rs_hol = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					If rs_hol.EOF Or rs_hol.BOF Then
						d = d
					Else
						d = d -1
					End If

					com_date = DateAdd("d", 1, com_date)
					rs_hol.Close()
				Loop
				Set rs_hol = Nothing

				' 1/19 �߰�
				acpt_day = DateValue(Mid(as_request_date, 1, 10))
				ddd = DateDiff("d", acpt_day, curr_day)

				If d > ddd Then
					d = ddd
				End If
				' 1/19 �߰� end

				' 2012-02-06
				If d = 1 Then
					curr_hh = Int(DatePart("h", Now()))
					acpt_hh = Int(DatePart("h", as_request_date))

					If acpt_hh > 13 And curr_hh < 12 Then
						d = 0
					End If
				End If
				' 2012-02-06 end

				dd = d
				'if d > 2 and d < 7 then
				'	dd = 3
				'end if
				'if d > 6 then
				'	dd = 7
				If d > 4 Then
					dd = 5
				End If
			Else
			' ���� ��� ��
				dd = 0
			End If

			If dd = days Then
	%>
	<tr valign="middle" class="style11">
		<td width="88" height="20" class="style11">
			<div align="center"><%=as_acpt_date%></div>
		</td>
		<td width="57" height="20" class="style11">
			<div align="center" class="style25"><%=as_acpt_man%></div>
		</td>
		<td width="56" height="20" class="style11">
			<div align="center" class="style25"><%=as_acpt_user%></div>
		</td>
		<td width="101" height="20" class="style11">
			<div align="center">
				<%=Replace(as_tel_ddd, " ", "")%>-<%=Replace(as_tel_no1, " ", "")%>-<%=Replace(as_tel_no2, " ", "")%>
			</div>
		</td>
		<td width="102" height="20" class="style11"><div align="center"><%=as_company%></div></td>
		<td width="101" height="20" class="style11"><div align="center"><%=as_dept%></div></td>
		<td width="165" height="20">
			<div align="center"><%=as_sido%>&nbsp;<%=as_gugun%>&nbsp;<%=as_dong%>&nbsp;<%=as_addr%></div>
		</td>
		<td width="63"><div align="center"><%=as_mg_ce%></div></td>
		<td width="110"><div align="center"><%=as_memo%></div></td>
		<td width="64"><div align="center"><%=as_request_date%></div></td>
		<td width="57"><div align="center"><%=as_request_time%></div></td>
		<td width="56"><div align="center"><%=as_as_type%></div></td>
		<td width="38"><div align="center"><%=as_process%></div></td>
		<td width="55"><div align="center"><%=into_reason%></div></td>
	</tr>
<%
			End If
		Next
	End If
%>
</table>
</body>
</html>