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
Dim com_tab
Dim com_sum(7)
Dim ok_sum(7)
Dim mi_sum(7)
Dim com_cnt(7,10)
Dim com_in(7,10)
Dim sum_cnt(10)
Dim sum_in(10)
Dim company_tab(150)
Dim end_tab(11)
Dim mi_tab(11)
Dim curr_mi_tab(11)

Dim i, j, k, l
Dim curr_day, curr_date, to_date, as_type, company, tot_cnt
Dim dd, a, d, com_date
Dim curr_hh, title_line
Dim asRs, weekRs
Dim sido, strSql, whereSql, groupSql
Dim rs_wek, holRs
Dim totSumCnt

title_line = "�湮ó�� ���纰 ��ó�� ��Ȳ (��û�� ����)"

com_tab = Array("����", "�λ�����", "�뱸����", "��������", "��������", "��������", "��������", "��������")

For i = 0 To 7
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0

	For j = 0 To 10
		com_cnt(i,j) = 0
		com_in(i,j) = 0
		sum_cnt(j) = 0
		sum_in(j) = 0
	Next
Next

For i = 0 To 11
	end_tab(i) = 0
	mi_tab(i) = 0
	curr_mi_tab(i) = 0
Next

curr_day = DateValue(Mid(CStr(Now()), 1, 10))	'���� ����
curr_date = DateValue(Mid(DateAdd("h", 12, Now()), 1, 10))	'���� ���� + 12�ð�

to_date = Mid(CStr(Now()), 1, 10)	'���� ����(curr_day �ߺ� ���ǵ�)

as_type = "�湮ó��"
company = "��ü"
mg_group = "1"



tot_cnt = 0

strSql = "as_process, CAST(request_date AS DATE) AS acpt_day, "
strSql = strSql & "CAST((request_date + INTERVAL 10 DAY_HOUR) AS DATE) AS com_date, "
strSql = strSql & "COUNT(*) AS err_cnt "
strSql = strSql & "FROM as_acpt "

whereSql = "WHERE as_type ='�湮ó��' AND mg_group ='1' "
whereSql = whereSql & "AND (as_process = '����' OR as_process = '�԰�' OR as_process = '����') "
whereSql = whereSql & "AND CAST(request_date AS DATE) <= NOW() "

groupSql = "GROUP BY sido, as_process, CAST(request_date AS DATE), "
groupSql = groupSql & "CAST((request_date + INTERVAL 10 DAY_HOUR) AS DATE) "

' �����õ�ÿ� �ܾ籺�� �������翡�� ��������� ������ ����� (2018-11-16)  ����� ���� �䱸
objBuilder.Append "SELECT sido, com_date, acpt_day, as_process, err_cnt "
objBuilder.Append "FROM ("
objBuilder.Append "	SELECT sido, " & strSql & whereSql
objBuilder.Append "		AND (sido <> '���' AND sido <> '����') "
objBuilder.Append groupSql & "UNION ALL "

objBuilder.Append "	SELECT '���', " & strSql & whereSql
objBuilder.Append "		AND (sido = '���' AND (gugun <> '��õ��' AND gugun <> '�ܾ籺')) "
objBuilder.Append groupSql & "UNION ALL "

objBuilder.Append "	SELECT '����', " & strSql & whereSql
objBuilder.Append "		AND (sido = '����' or (gugun = '��õ��' or gugun = '�ܾ籺')) "
objBuilder.Append groupSql
objBuilder.Append ") r ORDER BY sido ASC "

Set asRs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not asRs.EOF Then
	arrRs = asRs.getRows()
End If
asRs.Close() : Set asRs = Nothing

Dim arrRs
Dim as_sido, as_com_date, as_acpt_day, as_as_process, as_err_cnt

If IsArray(arrRs) Then
	For l = LBound(arrRs) To UBound(arrRs, 2)
		as_sido = arrRs(0, l)
		as_com_date = arrRs(1, l)
		as_acpt_day = arrRs(2, l)
		as_as_process = arrRs(3, l)
		as_err_cnt = arrRs(4, l)

		Select Case as_sido
			Case "����" : i = 0
			Case "���" : i = 0
			Case "��õ" : i = 0
			Case "�λ�" : i = 1
			Case "���" : i = 1
			Case "�泲" : i = 1
			Case "�뱸" : i = 2
			Case "���" : i = 2
			Case "����" : i = 3
			Case "�泲" : i = 3
			Case "���" : i = 3
			Case "����" : i = 3
			Case "����" : i = 4
			Case "����" : i = 4
			Case "����" : i = 4 ' 5 -> 4  ������ ��������� ���� (2018.09.27 ����)
			Case "����" : i = 6
			Case "����" : i = 7
		End Select

		'ó�� ��û���� ���� ���� : ��û�� + 10�ð��� �������� + 12�ð� ����
		dd = DateDiff("d", as_com_date, curr_date)

		If dd < 0 Then
			dd = 0
		End If

		'���� ���ڿ� ��û�� ��
		If CStr(curr_day) = CStr(as_acpt_day) Then
			dd = 0
		End If

		'���� ���
		If dd > 0 Then
			a = DateDiff("d", as_acpt_day, curr_day)
			'b = datepart("w",rs("acpt_day"))
			'bb = datepart("w", curr_day)
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

			com_date = DateValue(as_acpt_day)

			Do Until com_date > curr_day
				'sql_hol = "select * from (select DAYOFWEEK('" + cstr(com_date) + "') as  dayweeks ) A where A.dayweeks in (1,7)"
				objBuilder.Append "SELECT dayweeks FROM (SELECT DAYOFWEEK('" & CStr(com_date) & "') AS dayweeks) r "
				objBuilder.Append "WHERE r.dayweeks IN (1, 7); "

				Set weekRs = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If weekRs.EOF Or weekRs.BOF Then
					d = d
				Else
					d = d - 1
				End If

				com_date = DateAdd("d", 1, com_date)

				weekRs.Close()
			Loop
			Set weekRs = Nothing

			com_date = DateValue(as_acpt_day)

			Do Until com_date > curr_day
				objBuilder.Append "SELECT holiday FROM holiday WHERE holiday = '" & CStr(com_date) & "' "

				Set holRs = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If holRs.EOF Or holRs.BOF Then
					d = d
				Else
					d = d -1
				End If

				com_date = DateAdd("d", 1, com_date)

				holRs.Close()
			Loop
			Set holRs = Nothing

			' 2012-02-06
			If d = 1 Then
				curr_hh = Int(DatePart("h", Now()))

				If as_acpt_day <> as_com_date And curr_hh < 12 Then
					d = 0
				End If
			End If

			' 2012-02-06 end
			If d = 0 Then  '����
				j = 5
			ElseIf d = 1 Then  '����
				j = 6
			ElseIf d = 2 Then  '2��
				j = 7
	'		elseif d > 2 and d < 7  then
	'			j = 8
	'		else
	'			j = 9
			ElseIf d = 3 Then '3��
				j = 8
			ElseIf d = 4 Then '4��
				j = 9
			Else '5���̻�
				j = 10
			End If

			com_cnt(i, j) = com_cnt(i, j) + CLng(as_err_cnt)

			If as_as_process = "�԰�" Then
				com_in(i, j) = com_in(i, j) + CLng(as_err_cnt)
			End If
		  Else
	' ���� ��� ��
			com_cnt(i, 5) = com_cnt(i, 5) + CLng(as_err_cnt)
			'com_cnt(i,6) = com_cnt(i,6) + clng(rs("err_cnt"))

			If as_as_process = "�԰�" Then
				com_in(i, 5) = com_in(i, 5) + CLng(as_err_cnt)
				'com_in(i,6) = com_in(i,6) + clng(rs("err_cnt"))
			End If
		End If

		tot_cnt = tot_cnt + CLng(as_err_cnt)
	Next
End If
DBConn.Close() : Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>A/S ���� �ý���</title>
	<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
	<link href="/include/style.css" type="text/css" rel="stylesheet">

	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
	  function setCookie(cname, cvalue, exdays){
		  var d = new Date();
		  d.setTime(d.getTime() + (exdays*24*60*60*1000));

		  var expires = "expires="+ d.toUTCString();
		  document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
	  }

	  // '���ø� �� â�� ���� ����' Ŭ��
	  function closePop(){
		setCookie('first_as_view', 'first_as_view', 1);
		self.close();
	  }
	</script>
</head>
<body>
<div id="container">
	<h3 class="tit"><%=title_line%></h3>
	<form action="" method="post" name="frm">
		<div class="gView" style="position: relative;">
			<h3 class="stit">* ����ð� : <%=now()%></h3>
			<table cellpadding="0" cellspacing="0" class="tableList3" style="width:1000px;">
				<colgroup>
					<col width="*" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="6%" >
					<col width="5%" >
					<col width="10%" >
				</colgroup>
				<thead>
					<tr>
					  <th rowspan="2" class="first" scope="col">����</th>

						<th colspan="2" style=" border-left:1px solid #e3e3e3;border-bottom:1px solid #e3e3e3;" scope="col">����</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">����</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">2��</th>
						<!--
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3��~6��</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">7���̻�</th>
						-->
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3��</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">4��</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">5���̻�</th>
						<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">�Ұ�</th>
						<th rowspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">�����</th>
					</tr>
					<tr>
					  <th scope="col" style=" border-left:1px solid #e3e3e3;">�Ǽ�</th>
					  <th scope="col" style=" border-left:1px solid #e3e3e3;">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�Ǽ�</th>
					  <th style=" border-left:1px solid #e3e3e3;" scope="col">�԰�</th>
				  </tr>
				</thead>
				<tbody>
				<%
				If tot_cnt > 0 Then
					k = 0
				Else
					k = 7
				End If

'--------------------------------------���� Ȯ��
				For i = k To 7
					If com_tab(i) <> "" Then

						For j = 0 To 4
							ok_sum(i) = ok_sum(i) + com_cnt(i,j)
							sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)
						Next

						'for j = 5 to 9
						For j = 5 To 10
							mi_sum(i) = mi_sum(i) + com_cnt(i,j)
							sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)
							sum_in(j) = sum_in(j) + com_in(i,j)
						Next
						com_sum(i) = ok_sum(i) + mi_sum(i)

						sido = com_tab(i)
					End If
				Next
'--------------------------------------���� Ȯ��
				%>
					<tr>
					  <th>��</th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(5)), 0)%></a></th>
					  <th class="right"><%=sum_in(5)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(6)), 0)%></a></th>
					  <th class="right"><%=sum_in(6)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(7)), 0)%></a></th>
					  <th class="right"><%=sum_in(7)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(8)), 0)%></a></th>
					  <th class="right"><%=sum_in(8)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(9)), 0)%></a></th>
					  <th class="right"><%=sum_in(9)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(10)),0)%></a></th>
					  <th class="right"><%=sum_in(10)%></th>
					  <th class="right"><%=FormatNumber(CLng(sum_cnt(5) + sum_cnt(6) + sum_cnt(7) + sum_cnt(8) + sum_cnt(9) + sum_cnt(10)), 0)%></th>
					  <th class="right"><%=sum_in(5)+sum_in(6)+sum_in(7)+sum_in(8)+sum_in(9)+sum_in(10)%></th>
					  <th class="right">
					  <%
					  If tot_cnt = 0 Then
							Response.Write "0%"
					  Else
							totSumCnt = sum_cnt(0) + sum_cnt(1) + sum_cnt(2) + sum_cnt(3) + sum_cnt(4) + sum_cnt(5)
							totSumCnt = totSumCnt + sum_cnt(6) + sum_cnt(7) + sum_cnt(8) + sum_cnt(9) + sum_cnt(10)

							Response.Write FormatNumber(totSumCnt /tot_cnt * 100, 2) & "%"

							'=FormatNumber(((sum_cnt(0) + sum_cnt(1) + sum_cnt(2) + sum_cnt(3) + sum_cnt(4) + sum_cnt(5) + sum_cnt(6) + sum_cnt(7) + sum_cnt(8) + sum_cnt(9) + sum_cnt(10)) / tot_cnt * 100), 2)%
					  End If
					  %>
					  &nbsp;
					  </th>
					</tr>
				<%
				If tot_cnt > 0 Then
					k = 0
				Else
					k = 7
				End If

				For i = k To 7
					If com_tab(i) <> "" Then
					  ' �������簡 ������ (2018.09.27 ����)
						If i <> 5 Then
				%>
					<tr>
						<td><%=com_tab(i)%></td>
						<td class="right">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=0','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=FormatNumber(CLng(com_cnt(i, 5)), 0)%></a>
						</td>
						<td class="right"><%=com_in(i,5)%></td>
						<td class="right">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=FormatNumber(CLng(com_cnt(i, 6)), 0)%></a>
						</td>
						<td class="right"><%=com_in(i,6)%></td>
						<td class="right" bgcolor="#FFFF88">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i, 7)), 0)%></strong></a>
						</td>
						<td class="right"><strong><%=com_in(i,7)%></strong></td>
						<td class="right" bgcolor="#FFBE7D">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i, 8)), 0)%></strong></a>
						</td>
						<td class="right"><strong><%=com_in(i,8)%></strong></td>
						<td class="right" bgcolor="#FF8080">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=4%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i, 9)), 0)%></strong></a>
						</td>
						<td class="right"><strong><%=com_in(i,9)%></strong></td>

						<!-- �߰� 4�� -->
						<td class="right" bgcolor="#FF8080">
							<a href="#" onClick="pop_Window('/main/day_michulri_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=5%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=formatnumber(clng(com_cnt(i,10)),0)%></strong></a>
						</td>
						<td class="right"><strong><%=com_in(i,10)%></strong></td>
						<!-- �߰� 4�� -->

						<td class="right">
							<a href="#" onClick="pop_Window('/main/as_michulri_popup_request.asp?from_date=1900-01-01&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>','as_mi_popup','scrollbars=yes,width=800,height=600')">
								<%=formatnumber(clng(mi_sum(i)),0)%></a>
						</td>
						<td class="right"><%=com_in(i, 5) + com_in(i, 6) + com_in(i, 7) + com_in(i, 8) + com_in(i, 9) + com_in(i, 10)%></td>
						<td class="right">
						<%
						If tot_cnt = 0 Then
							Response.Write "0%"
						Else
							Response.Write FormatNumber((com_sum(i)/tot_cnt * 100),2) & "%"
						End If
						%>
						&nbsp;
						</td>
					</tr>
				<%
					End If ' �������簡 ������ (2018.09.27 ����)
				End If
				Next
				%>
				</tbody>
			</table>
		</div>
	</form>
</div>

�����õ�ÿ� �ܾ籺�� �������翡�� ��������� ����

<table cellpadding="0" cellspacing="0" style="width:1000px;">
	<tr>
		<td width="585" height="25" valign="middle">
			<div align="right">
				<span class="style1"><strong>���ø� �� â�� ���� ����</strong></span>
				<input name="todayPop" type="checkbox" id="todayPop" onClick="closePop();" value="checkbox">
			</div>
		</td>
	</tr>
</table>

</body>
</html>
