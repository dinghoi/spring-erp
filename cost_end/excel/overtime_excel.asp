<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim allowerIDs
Dim from_date, to_date, view_c, mg_ce, savefilename
Dim posi_sql, view_condi, view_grade, rsOT

' ��Ư�� ���α��� ID ����Ʈ
allowerIDs = Array("100125","100029","100015","100031","100020","100018") ' "����","�����","������","�ֱ漺','ȫ����','������'

from_date = Request("from_date")
to_date = Request("to_date")
view_c = Request("view_c")
mg_ce = Request("mg_ce")

savefilename = "��Ư�� ��Ȳ("&from_date&"_"&to_date&").xls"

'���� �ٿ�ε� ����
Call ViewExcelType(savefilename)

' �����Ǻ�
posi_sql = " AND mg_ce_id = '"&user_id&"'"

If position = "����" Then
	view_condi = "����"
End If

If position = "��Ʈ��" Then
	If view_c = "total" Then
		If org_name = "��ȭ����ȣ��" Then
			posi_sql = "AND (org_name = '��ȭ����ȣ��' OR org_name = '��ȭ��������') "
		Else
			posi_sql = "AND org_name = '"&org_name&"' "
		End If
	Else
		If org_name = "��ȭ����ȣ��" Then
			posi_sql = "AND (org_name = '��ȭ����ȣ��' OR org_name = '��ȭ��������') AND user_name LIKE '%"&mg_ce&"%' "
		Else
			posi_sql = "AND org_name = '"&org_name&"' AND user_name LIKE '%"&mg_ce&"%' "
		End If
	End If
End If

If position = "����" Then
	If view_c = "total" Then
		posi_sql = "AND ovrt.team = '"&team&"' "
	Else
		posi_sql = "AND ovrt.team = '"&team&"' AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

If position = "�������" Or cost_grade = "2" Then
	If view_c = "total" Then
        posi_sql = "AND ovrt.saupbu = emtt.emp_saupbu "
	Else
        posi_sql = "AND ovrt.saupbu = emtt.emp_saupbu AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

If position = "������" Or cost_grade = "1" Then
	If view_c = "total" Then
	  posi_sql = "AND ovrt.bonbu = '"&bonbu&"' "
	Else
	  posi_sql = "AND ovrt.bonbu = '"&bonbu&"' AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

view_grade = position

If cost_grade = "0" Then
	view_grade = "��ü"

  	If view_c = "total" Then
		posi_sql = ""
 	Else
		posi_sql = " AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

objBuilder.Append "SELECT ovrt.mg_ce_id, ovrt.work_date, ovrt.end_date, "
objBuilder.Append "	ovrt.from_time, ovrt.to_time, "
objBuilder.Append "	LEFT(ovrt.to_time, 2) AS totime,"
objBuilder.Append "	LEFT(ovrt.from_time, 2) AS fromtime, "
objBuilder.Append "	RIGHT(ovrt.to_time, 2) AS tominute, "
objBuilder.Append "	RIGHT(ovrt.from_time, 2) AS fromminute,"
objBuilder.Append "	ovrt.acpt_no, ovrt.user_name, ovrt.cost_detail,"
objBuilder.Append "	IFNULL(ovrt.delta_minute, 0) AS delta_minute,"
objBuilder.Append "	IFNULL(ovrt.rest_minute, 0) AS rest_minute,"
objBuilder.Append "	ovrt.alter_timeoff_date, ovrt.alter_timeoff_time, "
objBuilder.Append "	LEFT(ovrt.alter_timeoff_time, 2) AS altertimeofftime, "
objBuilder.Append "	RIGHT(ovrt.alter_timeoff_time, 2) AS altertimeoffminute, "
objBuilder.Append "	ovrt.alter_timeoff_minute_w, ovrt.alter_timeoff_minute_d, "
objBuilder.Append "	DATE_FORMAT(DATE_ADD(ovrt.alter_timeoff_date, "
objBuilder.Append "		INTERVAL(ovrt.alter_timeoff_minute_d) MINUTE), "
objBuilder.Append "		'%Y-%m-%d %I:%i') AS alter_timeoff_enddate1, "
objBuilder.Append "	DATE_FORMAT(DATE_ADD(ovrt.alter_timeoff_date, "
objBuilder.Append "	INTERVAL(ovrt.alter_timeoff_minute_w + ovrt.alter_timeoff_minute_d) MINUTE), "
objBuilder.Append "	'%Y-%m-%d %I:%i') AS alter_timeoff_enddate2, "
objBuilder.Append "	(SELECT visit_date FROM as_acpt WHERE acpt_no = ovrt.acpt_no) AS visit_date, "
objBuilder.Append "	ovrt.allow_yn, ovrt.allow_sayou, ovrt.cancel_yn, ovrt.you_yn, ovrt.reside_place, "
objBuilder.Append "	ovrt.user_name, ovrt.user_grade, ovrt.company, ovrt.dept, ovrt.cost_center, "
objBuilder.Append "	ovrt.work_gubun, ovrt.work_memo, ovrt.overtime_amt, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_name "
objBuilder.Append "FROM overtime AS ovrt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON ovrt.mg_ce_id = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE work_date BETWEEN '"&from_date&"' AND '"&to_date&"' "
objBuilder.Append posi_sql
objBuilder.Append "ORDER BY eomt.org_name, ovrt.user_name, ovrt.work_date "

'response.write objBuilder.ToString()
'Response.end
Set rsOT = Server.CreateObject("ADODB.RecordSet")
rsOT.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">���</th>
								<th scope="col">�۾���</th>
								<th scope="col">��Ư�� ����</th>
								<th scope="col">��Ư�� ��</th>
								<th scope="col">�ѽð�</th>
								<th scope="col">��ü�ް�</th>
								<th scope="col">AS NO</th>
								<th scope="col">ȸ��</th>
								<th scope="col">������</th>
								<th scope="col">�������</th>
								<th scope="col">��Ư�ٱ���</th>
								<th scope="col">�۾�����</th>
								<th scope="col">��û�ݾ�</th>
								<th scope="col">������</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">�̽��λ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim delta_minute, rest_minute, work_time, work_minute
						Dim cancel_yn, acpt_no, you_view, find, i
						Dim dateNow, week, mGap, fDate, lDate, rsChk
						Dim last_cnt

						Set rsChk = Server.CreateObject("ADODB.RecordSet")

						Do Until rsOT.EOF
						    delta_minute = CInt(rsOT("delta_minute")) ' �Ѱ���ð��� �Ѻ����� ..
                            rest_minute  = CInt(rsOT("rest_minute"))  ' ���ްԽð��� �Ѻ����� ..

                            If delta_minute > rest_minute Then
                                delta_minute = delta_minute - rest_minute
                            Else
                                delta_minute = 0
                            End If

                            work_time = Fix(delta_minute / 60) ' ���۾��ð��� �÷� ..
                            work_minute = delta_minute Mod 60    ' ���۾��ð��� �÷� �������� ������ ..

							If rsOT("cancel_yn") = "Y" Then
								cancel_yn = "���"
							Else
								cancel_yn = "����"
							End If

							If rsOT("acpt_no") = 0 Or rsOT("acpt_no") = null Then
								acpt_no = "����"
							Else
								acpt_no = rsOT("acpt_no")
							End If

							If rsOT("you_yn") = "Y" Then
								you_view = "����"
							Else
							 	you_view = "����"
							End If
                            %>
                            <tr>
                                <td class="first"><%=rsOT("org_company")%></td>
                                <td><%=rsOT("org_bonbu")%></td>
                                <td><%=rsOT("org_saupbu")%></td>
                                <td><%=rsOT("org_team")%></td>
                                <td><%=rsOT("org_name")%></td>
                                <td><%=rsOT("reside_place")%></td>
                                <td><%=rsOT("mg_ce_id")%></td>
                                <td><%=rsOT("user_name")%>&nbsp;<%=rsOT("user_grade")%></td>

                                <td><%=rsOT("work_date")%>&nbsp;<%=rsOT("fromtime")%>:<%=rsOT("fromminute")%></td>
                                <td><%=rsOT("end_date")%>&nbsp;<%=rsOT("totime")%>:<%=rsOT("tominute")%></td>
                                <td><%=work_time%>�ð� <%=work_minute%>��</td>
                                <td>
								<%
                                If rsOT("alter_timeoff_date") <> "" Then '����ڰ� ��ü�ް��������� �Է����� ���
                                %>
                                    <%=rsOT("alter_timeoff_date")%>&nbsp;<%=rsOT("altertimeofftime")%>:<%=rsOT("altertimeoffminute")%>
                                    <br> ~
                                    <%
                                    If CInt(rsOT("alter_timeoff_minute_w")) > 0 Then ' 52�ð� �ʰ����� ���
                                        dateNow = CDate(rsOT("work_date")) ' ���ں�ȯ
										week = Weekday(dateNow)	' ����

										If week >= 4 Then
											mGap = (week - 4) * -1
										Else
											mGap = (6 - (3 - week)) * -1
										End If

										fDate = DateAdd("d", mGap, dateNow)
										lDate = DateAdd("d", mGap + 6, dateNow)

										objBuilder.Append "SELECT COUNT(*) AS last_cnt "
										objBuilder.Append "FROM overtime "
										objBuilder.Append "WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"' "
										objBuilder.Append "	AND mg_ce_id  = '"&rsOT("mg_ce_id")&"' "
										objBuilder.Append "	AND LENGTH(alter_timeoff_date) > 0 "
										objBuilder.Append "	AND work_date > '"&rsOT("work_date")&"' "

										rsChk.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										last_cnt = 0

										If Not (rsChk.BOF Or rsChk.EOF) Then
											last_cnt = CInt(RsChk("last_cnt"))
										End If

										rsChk.Close()

										If last_cnt = 0 Then  ' ������ 52�ð� �ʰ����� ���
											Response.Write rsOT("alter_timeoff_enddate2") ' �� 52�ð� �ʰ� + (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
										Else
											Response.Write rsOT("alter_timeoff_enddate1") ' (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
										End If
                                    Else ' 52�ð� �ʰ����� �ƴ� ���
										Response.write rsOT("alter_timeoff_enddate1") ' (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
                                    End If
                                End If
                                %>
								</td>
								<td><%=acpt_no%></td>
								<td><%=rsOT("company")%></td>
								<td><%=rsOT("dept")%></td>
								<td><%=rsOT("cost_center")%></td>
								<td><%=rsOT("work_gubun")%></td>
								<td><%=rsOT("work_memo")%></td>
								<%
  								find = False

                                For i = 0 To UBound(allowerIDs)
                                    if  user_id = allowerIDs(i) then
                                        find =True
                                    end if
                                Next

                                If find = True Then
                                %>
								<td class="right"><%=formatnumber(rsOT("overtime_amt"),0)%></td>
								<%
                                End If
  							    %>
								<td><%=you_view%></td>
                                <td><%=cancel_yn%></td>
								<td><%=rsOT("allow_yn")%></td>
								<td><span name ="allowSayou"><%=rsOT("allow_sayou")%></span></td>
							</tr>
						    <%
							rsOT.MoveNext()
						Loop
						rsOT.Close() : Set rsOT = Nothing
						Set rsChk = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>