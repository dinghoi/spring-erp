<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim run_month, transit_id, view_c, view_d, use_man
Dim from_date, end_date, to_date, sign_month, savefilename
Dim posi_sql, view_grade, transit_sql, base_sql, date_sql, order_sql
Dim rsTran, arrTran

run_month = Request.QueryString("run_month")
transit_id = Request.QueryString("transit_id")
view_c = Request.QueryString("view_c")
view_d = Request.QueryString("view_d")
use_man = Request.QueryString("use_man")

If run_month = "" Then
	run_month = Mid(CStr(Now()),1,4)&Mid(CStr(Now()),6,2)
	view_c = "total"
	emp_name = ""
End If

from_date = Mid(run_month,1,4)&"-"&Mid(run_month,5,2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d",-1,end_date))
sign_month = run_month

savefilename = run_month& "�� "&transit_id&" ����� ��Ȳ.xls"

' �����Ǻ�
'posi_sql = "AND transit_cost.mg_ce_id = '"&user_id&"' "
posi_sql = "AND trct.mg_ce_id = '"&user_id&"' "

'"��ȭ���� ����"�� ��� "��ȭ���� ����" ���絵 Ȯ�� �����ϰ� �߰�(������ �븮 ��û)[����ȣ_20210809]
If position = "��Ʈ��" Then
	Select Case org_name
		Case "��ȭ���� ȣ��"
			posi_sql = "AND (trct.org_name = '��ȭ���� ȣ��' OR trct.org_name = '��ȭ���� ����') "
		Case "��ȭ���� ����"
			posi_sql = "AND (trct.org_name = '"&org_name&"' OR trct.org_name = '��ȭ���� ����') "
		Case Else
			posi_sql = "AND trct.org_name = '"&org_name&"' "
	End Select

	If view_c <> "total" Then
		posi_sql = posi_sql&"AND memt.user_name LIKE '%"&use_man&"%' "
	End If
End If

If position = "����" Then
	posi_sql = "AND trct.team = '"&team&"' "

	If view_c <> "total" Then
        posi_sql = posi_sql&"AND memt.user_name LIKE '%"&use_man&"%' "
	End If
End If

If position = "�������" Or cost_grade = "2" Then
	posi_sql = " AND trct.saupbu = emp_master.emp_saupbu "

    If view_c = "total" Then
        posi_sql = posi_sql&"AND memt.user_name LIKE '%"&use_man&"%' "
    End If
End If


If position = "������" Or cost_grade = "1" Then
	posi_sql = "AND trct.bonbu = '"&bonbu&"' "

  	If view_c = "total" Then
		posi_sql = posi_sql&"AND memt.user_name LIKE '%"&use_man&"%' "
	End If
End If

If cost_grade = "0" Then
  	If view_c = "total" Then
		posi_sql = ""
 	Else
		posi_sql = "AND memt.user_name LIKE '%"&use_man&"%'"
	End If
End If

If transit_id = "����" Then
	transit_sql = "AND (trct.car_owner = '����' OR trct.car_owner = 'ȸ��') "
Else
	transit_sql = "AND (trct.car_owner = '���߱���') "
End If

If view_d = "run" Then
    date_sql = "AND (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
    order_sql = "ORDER BY memt.user_name, run_date DESC, run_seq DESC "
End If

If view_d = "reg" Then
    date_sql = "AND (trct.reg_date >= '"&from_date&" 00:00:00' AND trct.reg_date <='"&to_date&" 23:59:59') "
    order_sql = "ORDER BY memt.user_name, trct.reg_date DESC, run_seq DESC "
End If

'���Ǻ� ��ȸ
objBuilder.Append "SELECT run_date, mg_ce_id, run_seq, trct.user_name, "
objBuilder.Append "	oil_kind, start_company, start_point, far, transit, "
objBuilder.Append "	car_owner, start_km, end_km, oil_price, "
objBuilder.Append "	fare, run_memo, repair_cost, parking, toll, cancel_yn, "
objBuilder.Append "	end_yn, trct.reg_date, end_company, end_point, "
objBuilder.Append "	trct.emp_company, trct.bonbu, trct.saupbu, trct.team, trct.org_name, "
objBuilder.Append "	trct.reside_place, trct.company, trct.user_name, trct.cost_center "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id AND memt.grade < '5' "
objBuilder.Append "INNER JOIN emp_master AS emtt ON memt.user_id = emtt.emp_no "
objBuilder.Append "WHERE 1=1 "
objBuilder.Append transit_sql&posi_sql&date_sql&order_sql

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTran.EOF Then
	arrTran = rsTran.getRows()
End If
rsTran.Close() : Set rsTran = Nothing
DBConn.Close() : Set DBConn = Nothing

If IsArray(arrTran) Then
	'// ������ ����
	Call ViewExcelType(savefilename)
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
								<th rowspan="2" scope="col" class="first">ȸ��</th>
								<th rowspan="2" scope="col">����</th>
								<th rowspan="2" scope="col">�����</th>
								<th rowspan="2" scope="col">��</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">����ó</th>
								<th rowspan="2" scope="col">���ȸ��</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">���</th>
								<th rowspan="2" scope="col">��������</th>
								<th rowspan="2" scope="col">�߱�����</th>
								<th rowspan="2" scope="col">�������</th>
								<th rowspan="2" scope="col">��������</th>
								<th rowspan="2" scope="col">�����</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">�������</th>
								<th rowspan="2" scope="col">����KM</th>
								<th rowspan="2" scope="col">����KM</th>
								<th rowspan="2" scope="col">�Ÿ�</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� �� </th>
								<th rowspan="2" scope="col">����</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">������</th>
								<th scope="col">���߱���</th>
								<th scope="col">�����ݾ�</th>
								<th scope="col">������</th>
								<th scope="col">�����</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, run_date, mg_ce_id, run_seq, t_user_name, oil_kind, start_company, start_point
						Dim far, transit, car_owner, start_km, end_km, oil_price, fare, run_memo, repair_cost
						Dim parking, toll, cancel_yn, end_yn, reg_date, car_gubun, run_km, cancel_view
						Dim start_view, end_view, chk_slip_month, chk_reg_month, chk_reg_day, bgcolor
						Dim end_company, end_point, emp_bonbu, emp_saupbu, emp_team
						Dim trade_company, emp_name, cost_center

						For i=LBound(arrTran) To UBound(arrTran, 2)
							run_date = arrTran(0, i)
							mg_ce_id = arrTran(1, i)
							run_seq = arrTran(2, i)
							t_user_name = arrTran(3, i)
							oil_kind = arrTran(4, i)
							start_company = arrTran(5, i)
							start_point = arrTran(6, i)
							far = arrTran(7, i)
							transit = arrTran(8, i)
							car_owner = arrTran(9, i)
							start_km = arrTran(10, i)
							end_km = arrTran(11, i)
							oil_price = arrTran(12, i)
							fare = arrTran(13, i)
							run_memo = arrTran(14, i)
							repair_cost = arrTran(15, i)
							parking = arrTran(16, i)
							toll = arrTran(17, i)
							cancel_yn = arrTran(18, i)
							end_yn = arrTran(19, i)
							reg_date = arrTran(20, i)
							end_company = arrTran(21, i)
							end_point = arrTran(22, i)
							emp_company = arrTran(23, i)
							emp_bonbu = arrTran(24, i)
							emp_saupbu = arrTran(25, i)
							emp_team = arrTran(26, i)
							org_name = arrTran(27, i)
							reside_place = arrTran(28, i)
							trade_company = arrTran(29, i)
							emp_name = arrTran(30, i)
							cost_center = arrTran(31, i)

							If car_owner = "���߱���" Then
								car_gubun = transit
							Else
								car_gubun = car_owner&"<br>"&oil_kind
							End If

							run_km = far

							If cancel_yn = "Y" Then
								cancel_view = "���"
							Else
								cancel_view = "����"
							End If

							If f_toString(start_km, "") = "" Then
								start_view = 0
							Else
								start_view = start_km
							End If

							If f_toString(end_km, "") = "" Then
								end_view = 0
							Else
								end_view = end_km
							End If

                            ' 5�� ���� ���� �Է°� ����
							chk_slip_month = Mid(run_date,1,7)
							chk_reg_month = Mid(reg_date,1,7)
							chk_reg_day = Mid(reg_date,9,2)

							If chk_slip_month < chk_reg_month And chk_reg_day > "05" Then
								bgcolor = "burlywood"
							Else
								bgcolor = "#f8f8f8"
							End If
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
                                <td class="first"><%=emp_company%></td>
                                <td><%=emp_bonbu%></td>
                                <td><%=emp_saupbu%></td>
                                <td><%=emp_team%></td>
                                <td><%=org_name%></td>
                                <td><%=reside_place%></td>
                                <td><%=trade_company%></td>
                                <td><%=emp_name%></td>
                                <td><%=mg_ce_id%></td>
                                <td><%=run_date%></td>
                                <td><%=Mid(reg_date,1,10)%></td>
                                <td><%=cost_center%></td>
                                <td><%=car_gubun%></td>
                                <td><%=start_company%>-<%=start_point%></td>
                                <td><%=end_company%>-<%=end_point%></td>
                                <td><%=run_memo%>&nbsp;</td>
                                <td class="right"><%=formatnumber(start_view,0)%></td>
                                <td class="right"><%=formatnumber(end_view,0)%></td>
                                <td class="right"><%=formatnumber(run_km,0)%></td>
                                <td class="right"><%=formatnumber(repair_cost,0)%></td>
                                <td class="right"><%=formatnumber(fare,0)%></td>
                                <td class="right"><%=formatnumber(oil_price,0)%></td>
                                <td class="right"><%=formatnumber(parking,0)%></td>
                                <td class="right"><%=formatnumber(toll,0)%></td>
                                <td><%=cancel_view%></td>
                            </tr>
                        <%
						Next
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>
<%
Else
	Response.Write "<script>alert('�����Ͱ� �������� �ʽ��ϴ�.');history.go(-1);</script>"
	Response.End
End If
%>