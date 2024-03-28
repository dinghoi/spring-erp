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
Dim company_tab(50)
Dim from_date, to_date, sido, mg_ce, mg_ce_id, title_line
Dim company, as_type, win_sw, memo01, memo02, i
Dim type_sql, in_cnt, acpt_cnt, yun_cnt, grade_sql, com_sql
Dim rsAs, arrAs

from_date = f_Request("from_date")
to_date = f_Request("to_date")
sido = f_Request("sido")
mg_ce = f_Request("mg_ce")
mg_ce_id = f_Request("mg_ce_id")
mg_group = f_Request("mg_group")
company = f_Request("company")
as_type = f_Request("as_type")

title_line = "��ó�� ��Ȳ"
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

If as_type = "��ü" Then
	type_sql = ""
Else
  	type_sql = "as_type ='"&as_type&"' AND "
End If

i = 0
in_cnt = 0
acpt_cnt = 0
yun_cnt = 0

If company = "��ü" And c_grade = "7" Then
	k = 0

	'Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and group_name = '"+user_name+"' order by etc_name asc"
	objBuilder.Append "SELECT etc_name FROM etc_code WHERE etc_type = '51' AND used_sw = 'Y' AND group_name = '"&user_name&"' "
	objBuilder.Append "ORDER BY etc_name ASC "

	Set rs_etc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	While Not rs_etc.EOF
		k = k + 1

		company_tab(k) = rs_etc("etc_name")
		rs_etc.MoveNext()
	Wend
	rs_etc.close() : Set rs_etc = nothing
End If

If company = "��ü" Then
	grade_sql = ""
Else
	grade_sql = "company = '"&company&"' AND "
End If

If c_grade = "7" And company = "��ü" Then
	com_sql = "company = '"&company_tab(1)&"' "

	For kk = 2 To k
		com_sql = com_sql & " OR company = '"&company_tab(kk)&"' "
	Next

	grade_sql = "("&com_sql&") AND "
End If

If (c_grade = "8") Or (c_grade = "7" And company <> "��ü") Then
	grade_sql = "(company = '"&company&"') AND "
End If

com_sql = grade_sql

objBuilder.Append "SELECT acpt_no, request_date, as_process, company, dept, sido, gugun, as_type, acpt_date "
objBuilder.Append "FROM as_acpt "
objBuilder.Append "WHERE "&com_sql&type_sql&" (as_process = '����' or as_process = '�԰�' or as_process = '����') "
objBuilder.Append "	AND (CAST(request_date as date) >= '"&from_date&"' AND CAST(request_date as date) <= '"&to_date&"') "

If mg_ce = "" Then
	Select Case sido
		Case "��"
			objBuilder.Append ""
		Case "����"
			objBuilder.Append "	AND sido IN ('����', '���', '��õ') "
		Case "�λ�����"
			objBuilder.Append "	AND sido IN ('�λ�', '�泲', '���') "
		Case "�뱸����"
			objBuilder.Append "	AND sido IN ('�뱸', '���') "
		Case "��������"
			objBuilder.Append "	AND sido IN ('����', '�泲', '���', '����') "
		Case "��������"
			objBuilder.Append "	AND sido IN ('����', '����', '����') "
		Case "��������"
			objBuilder.Append "	AND sido  = '����' "
		Case "��������"
			objBuilder.Append "	AND sido  = '����' "
		Case Else
			objBuilder.Append "	AND sido = '"&sido&"' "
			objBuilder.Append " ORDER BY acpt_date ASC "
	End Select
  'if   sido = "��" then
    'sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    'sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
  'elseif sido = "����" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','���','��õ')"
  'elseif sido = "�λ�����" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�λ�','�泲','���')"
  'elseif sido = "�뱸����" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�뱸','���')"
  'elseif sido = "��������" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','�泲','���','����')"
  ' ������ ��������� ���� (2018.09.27 ����)
  'elseif sido = "��������" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','����','����')"
  ' �������簡 ������ (2018.09.27 ����)

  'elseif sido = "��������" then - ���� �ּ�
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  ' - ���� �ּ� ��

  'elseif sido = "��������" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  'elseif sido = "��������" then
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  'else
'		sql = "select * from as_acpt"
'		sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (sido = '" + sido + "')"
'		sql = sql + "  and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') Order By acpt_date Asc"
'	end if
Else
	'sql = "select * from as_acpt"
	'sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (mg_ce_id = '" + mg_ce_id + "')"
	'sql = sql + "  and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') Order By acpt_date Asc"

	objBuilder.Append "	AND mg_ce_id = '"&mg_ce_id&"' "
	objBuilder.Append "ORDER BY acpt_date ASC "
End If

If from_date = "" Then
	objBuilder.Clear()

	'sql = "select * from as_acpt"
	'sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (sido = '" + sido + "')"
	'sql = sql + " Order By acpt_date Asc"
	objBuilder.Append "SELECT acpt_no, request_date, as_process, company, dept, sido, gugun, as_type, acpt_date "
	objBuilder.Append "FROM as_acpt "
	objBuilder.Append "WHERE "&com_sql&type_sql&" (as_process = '����' or as_process = '�԰�' or as_process = '����') "
	objBuilder.Append "	AND sido = '"&sido&"' "
	objBuilder.Append "ORDER BY acpt_date ASC "
End If

Set rsAS = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAs.EOF Then
	arrAs = rsAs.getRows()
End If
rsAs.Close() : Set rsAS = Nothing
DBConn.Close() : Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ó�� ��Ȳ</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
		  		 window.close();
			}
        </script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="container">
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="10%" >
							<col width="12%" >
							<col width="10%" >
							<col width="20%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th><%=memo01%></th>
							  <td class="left"><%=memo02%></td>
							  <th>ȸ��</th>
							  <td class="left"><%=company%></td>
							  <th>ó������</th>
							  <td class="left"><%=as_type%></td>
							  <td>
								<a href = "as_michulri_excel_request.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=sido%>&company=<%=company%>&as_type=<%=as_type%>&mg_ce=<%=mg_ce%>&mg_ce_id=<%=mg_ce_id%>&mg_group=<%=mg_group%>" class="btnType04">�����ٿ�ε�</a>
							  </td>
					      	</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="5%" >
							<col width="18%" >
							<col width="25%" >
							<col width="*" >
							<col width="10%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��û����</th>
								<th scope="col">����</th>
								<th scope="col">ȸ���</th>
								<th scope="col">�μ���</th>
								<th scope="col">����</th>
								<th scope="col">ó������</th>
								<th scope="col">��ȸ</th>
							</tr>
						</thead>
						<tbody>
					  	<%
						'Dim int date_len
						Dim l, acpt_no, request_date, as_process, as_company, as_dept
						Dim as_sido, as_gugun, as_as_type, acpt_date, date_to_date
						Dim len_date, hangle, bit01, bit02, bit03

                        If IsArray(arrAs) Then
							For l = LBound(arrAs) To UBound(arrAs)
								acpt_no = arrAs(0, l)
								request_date = arrAs(1, l)
								as_process = arrAs(2, l)
								as_company = arrAs(3, l)
								as_dept = arrAs(4, l)
								as_sido = arrAs(5, l)
								as_gugun = arrAs(6, l)
								as_as_type = arrAs(7, l)
								acpt_date = arrAs(8, l)

								'date_len=len(rs("acpt_date"))
								'acpt_date = rs("acpt_date")

								len_date = Len(acpt_date)
								bit01 = Left(acpt_date, 10)
								'bit01 = Replace(bit01,"-",".")
								bit03 = Left(Right(acpt_date, 5), 2)
								hangle = Mid(acpt_date, 12, 2)

								If len_date = 22 Then
									bit02 = Mid(acpt_date, 15, 2)
								Else
									bit02 = "0"&Mid(acpt_date, 15, 1)
								End If

								If hangle = "����" And bit02 <> 12 Then
									bit02 = bit02 + 12
								End If

								date_to_date = bit01 & " " &bit02 & ":" & bit03
								acpt_date = Mid(date_to_date, 3)
								acpt_date = Replace(acpt_date, "-", "/")
								'acpt_date = rs("request_date")
								acpt_date = request_date

								If as_process = "����" Then
									acpt_cnt = acpt_cnt + 1
								End If

								If as_process = "����" Then
									yun_cnt = yun_cnt + 1
								End If

								If as_process = "�԰�" Then
									in_cnt = in_cnt + 1
								End If

								i = i + 1
                      	%>
						<tr>
							<td class="first"><%=i%></td>
							<td><%=acpt_date%></td>
							<td><%=as_process%></td>
							<td><%=as_company%></td>
							<td><%=as_dept%></td>
							<td><%=as_sido%>&nbsp;<%=as_gugun%></td>
							<td><%=as_as_type%></td>
							<td><a href="#" onClick="pop_Window('/main/as_view.asp?acpt_no=<%=acpt_no%>&win_sw=<%=win_sw%>','asview_pop','scrollbars=yes,width=800,height=700')">��ȸ</a></td>
						</tr>
						<%
							Next
						End If
						%>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>����</th>
							  <td class="left"><%=acpt_cnt%></td>
							  <th>����</th>
							  <td class="left"><%=yun_cnt%></td>
							  <th>�԰�</th>
							  <td class="left"><%=in_cnt%></td>
					      	</tr>
						</tbody>
					</table>
					<br>
				</form>
				</div>
			</div>
	</body>
</html>