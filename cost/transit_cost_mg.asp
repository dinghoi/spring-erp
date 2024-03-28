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
Dim run_month, transit_id, view_c, view_d, use_man, page
Dim from_date, end_date, to_date, sign_month
Dim pgsize, start_page, stpage, be_pg, str_param, total_page, total_record
Dim posi_sql, view_condi, view_grade, transit_sql, date_sql, order_sql
Dim rsCount, rs_sum, rsTran, title_line, arrTran
Dim sum_far, sum_oil_price, sum_fare, sum_repair_cost, sum_parking, sum_toll

run_month = f_Request("run_month")
transit_id = f_Request("transit_id")
view_c = f_Request("view_c")
view_d = f_Request("view_d")
use_man = f_Request("use_man")
page = f_Request("page")

title_line = "����� ����"

If view_d = "" Then
    view_d = "run"
End If

If run_month = "" Then
	run_month = Mid(CStr(Now()),1,4)&Mid(CStr(Now()),6,2)
	transit_id = "����"
    view_c = "total"
    view_d = "run"
	use_man = ""
End If

from_date = Mid(run_month,1,4)&"-"&Mid(run_month,5,2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d",-1,end_date))
sign_month = run_month

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If
stpage = Int((page-1)*pgsize)

str_param= "&run_month="&run_month&"&view_c="&view_c&"&view_d="&view_d&"&use_man="&use_man&"&transit_id="&transit_id

' �����Ǻ�
posi_sql = "AND trct.mg_ce_id = '"&user_id&"' "

If position = "����" Then
	view_condi = "����"
End If

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

view_grade = position

If cost_grade = "0" Then
	view_grade = "��ü"

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

'��ü ī��Ʈ
objBuilder.Append "SELECT COUNT(*) FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id AND memt.grade < '5' "
objBuilder.Append "INNER JOIN emp_master AS emtt ON memt.user_id = emtt.emp_no "
objBuilder.Append "WHERE 1=1 "&transit_sql&posi_sql&date_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0))'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record/pgsize) 'Result.PageCount
Else
	total_page = Int((total_record/pgsize)+1)
End If

objBuilder.Append "SELECT SUM(far) AS 'far', SUM(oil_price) AS 'oil_price', SUM(fare) AS 'fare', "
objBuilder.Append "	SUM(repair_cost) AS 'repair_cost', SUM(parking) AS 'parking', SUM(toll) AS 'toll'"
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id AND memt.grade < '5' "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emtt.emp_no = memt.user_id "
objBuilder.Append "WHERE cancel_yn = 'N' "&transit_sql&posi_sql&date_sql

Set rs_sum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rs_sum("far"), "") = "" Then
	sum_far = 0
	sum_oil_price = 0
	sum_fare = 0
	sum_repair_cost = 0
	sum_parking = 0
	sum_toll = 0
Else
	sum_far = rs_sum("far")
	sum_oil_price = rs_sum("oil_price")
	sum_fare = rs_sum("fare")
	sum_repair_cost = rs_sum("repair_cost")
	sum_parking = rs_sum("parking")
	sum_toll = rs_sum("toll")
End If

rs_sum.Close() : Set rs_sum = Nothing

'���Ǻ� ��ȸ
objBuilder.Append "SELECT run_date, mg_ce_id, run_seq, trct.user_name, "
objBuilder.Append "	oil_kind, start_company, start_point, far, transit, "
objBuilder.Append "	car_owner, start_km, end_km, oil_price, "
objBuilder.Append "	fare, run_memo, repair_cost, parking, toll, cancel_yn, "
objBuilder.Append "	end_yn, trct.reg_date, end_company, end_point "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id AND memt.grade < '5' "
objBuilder.Append "INNER JOIN emp_master AS emtt ON memt.user_id = emtt.emp_no "
objBuilder.Append "WHERE 1=1 "
objBuilder.Append transit_sql&posi_sql&date_sql&order_sql
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTran.EOF Then
	arrTran = rsTran.getRows()
End If
rsTran.Close() : Set rsTran = Nothing
DBConn.Close() : Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			$(document).ready(function(){
				condi_view();
			});

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.run_month.value == ""){
					alert ("�������� �Է��ϼ���");
					return false;
				}
				return true;
			}

			function condi_view(){
				var position = '<%=position%>';
				var cost_grade = '<%=cost_grade%>';

				if(position != '����' || cost_grade === '0'){
                    if(eval("document.frm.view_c[0].checked")){
                        document.getElementById('use_man_view').style.display = 'none';
                    }

                    if(eval("document.frm.view_c[1].checked")){
                        document.getElementById('use_man_view').style.display = '';
                    }
				}
			}

			function enterkey(){
				if(window.event.keyCode == '13'){
					frmcheck();
				}
				return;
			}
		</script>
	</head>
	<!--<body onLoad="condi_view();">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="transit_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
								<label>
                                    <input type="radio" name="view_d" value="run" <%If view_d = "run" Then %>checked<%End If %> style="width:20px"/>
                                    <strong>������&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <%If view_d = "reg" Then %>checked<%End If %> style="width:20px"/>
                                    <strong>�߱޳��&nbsp;</strong>

                                    : <input name="run_month" type="text" value="<%=run_month%>" style="width:70px" onkeyup="enterkey()"/>
                                    (��201401)
								</label>
								<label>
                              	<input type="radio" name="transit_id" value="����" <%If transit_id = "����" Then %>checked<%End If %> style="width:20px"/>
                                ������������
                                <input type="radio" name="transit_id" value="����" <%If transit_id = "����" Then %>checked<%End If %> style="width:20px"/>
                                ���߱����
								</label>
								<label><strong>��ȸ����:</strong><%=view_grade%></label>
								<label>
								<strong>��ȸ����:</strong>
								<%
								If position = "����" and cost_grade <> "0" Then
									Response.Write view_condi
                                Else
								%>
                                <input type="radio" name="view_c" value="total" <%If view_c = "total" Then %>checked <%End If %> style="width:20px" onClick="condi_view();"/>
                                    ������ü
                                <input type="radio" name="view_c" value="reg_id" <%If view_c = "reg_id" Then %>checked <%End If %> style="width:20px" onClick="condi_view();"/>
                                    ���κ�
								<%End If%>
                                </label>
								<label>
                                	<input name="use_man" type="text" value="<%=use_man%>" style="width:70px; display:none" id="use_man_view" onkeyup="enterkey()"/>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="17%" >
							<col width="17%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="3%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
							<col width="3%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">������</th>
                                <th rowspan="2" scope="col">��������</th>
                                <th rowspan="2" scope="col">�߱�����</th>
								<th rowspan="2" scope="col">����<br>����</th>
								<th rowspan="2" scope="col">�����</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">�������</th>
								<th rowspan="2" scope="col">����KM</th>
								<th rowspan="2" scope="col">����KM</th>
								<th rowspan="2" scope="col">�Ÿ�</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� �� </th>
								<th rowspan="2" scope="col">����</th>
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
						Dim end_company, end_point

						If IsArray(arrTran) Then
							For i = LBound(arrTran) To UBound(arrTran, 2)
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

								' 5�� ���� ���� �Է°� ����...
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
								<td class="first"><%=t_user_name%></td>
                                <td><%=run_date%></td>
                                <td><%=Mid(reg_date,1,10)%></td>
								<td><%=car_gubun%></td>
								<td><%=start_company%>-<%=start_point%></td>
								<td><%=end_company%>-<%=end_point%></td>
								<td><%=run_memo%>&nbsp;</td>
								<td class="right"><%=FormatNumber(start_view,0)%></td>
								<td class="right"><%=FormatNumber(end_view,0)%></td>
								<td class="right"><%=FormatNumber(run_km,0)%></td>
								<td class="right"><%=FormatNumber(repair_cost,0)%></td>
								<td class="right"><%=FormatNumber(fare,0)%></td>
								<td class="right"><%=FormatNumber(oil_price,0)%></td>
								<td class="right"><%=FormatNumber(parking,0)%></td>
								<td class="right"><%=FormatNumber(toll,0)%></td>
								<td><%=cancel_view%></td>
								<td>
								<%
								If end_yn <> "Y" Then
									If car_owner = "���߱���" Then
										If mg_ce_id = user_id Then
								%>
											<a href="#" onClick="pop_Window('/cost/mass_transit_add.asp?mg_ce_id=<%=mg_ce_id%>&run_date=<%=run_date%>&run_seq=<%=run_seq%>&u_type=U','mass_transit_add_pop','scrollbars=yes,width=850,height=350')">����</a>
								<%
										Else
								%>
											<a href="#" onClick="pop_Window('/cost/mass_transit_cancel.asp?mg_ce_id=<%=mg_ce_id%>&run_date=<%=run_date%>&run_seq=<%=run_seq%>&u_type=U','mass_transit_cancel_pop','scrollbars=yes,width=850,height=350')">����</a>
								<%
										End If
									Else
										If mg_ce_id = user_id Then
								%>
											<a href="#" onClick="pop_Window('/cost/car_drive_add.asp?mg_ce_id=<%=mg_ce_id%>&run_date=<%=run_date%>&run_seq=<%=run_seq%>&u_type=U','car_drive_add_pop','scrollbars=yes,width=900,height=500')">����</a>
								<%
										Else
								%>
											<a href="#" onClick="pop_Window('/cost/car_drive_cancel.asp?mg_ce_id=<%=mg_ce_id%>&run_date=<%=run_date%>&run_seq=<%=run_seq%>&u_type=U','car_drive_cancel_pop','scrollbars=yes,width=900,height=470')">����</a>
								<%		End If
									End If
								Else
									Response.Write "����"
								End If
								%>
                                </td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='17' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If

						If total_record <> 0 Then
						%>
							<tr>
								<th class="first">��</th>
								<th colspan="3"><%=total_record%>&nbsp;��</th>
								<th colspan="13">����Ÿ� :&nbsp;<%=FormatNumber(sum_far,0)%>&nbsp;KM&nbsp;&nbsp;,&nbsp;&nbsp;������ :&nbsp;<%=FormatNumber(sum_repair_cost,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;���߱���� :&nbsp;<%=FormatNumber(sum_fare,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;�����ݾ� :&nbsp;<%=FormatNumber(sum_oil_price,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;������ :&nbsp;<%=FormatNumber(sum_parking,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;����� :&nbsp;<%=FormatNumber(sum_toll,0)%></th>
							</tr>
						<%
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
                    <div class="btnCenter">
                        <a href="/cost/transit_cost_excel.asp?run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>" class="btnType04">�����ٿ�ε�</a>
                    </div>
                  	</td>
				    <td>
					<%
					'page navigator[����ȣ_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
                    </td>
				    <td width="25%">
                    <div class="btnCenter">
                        <a href="#" onClick="pop_Window('/cost/car_drive_add.asp','car_drive_add_popup','scrollbars=yes,width=900,height=450')" class="btnType04">�������������Է�</a>
                        <a href="#" onClick="pop_Window('/cost/mass_transit_add.asp','mass_transit_add_popup','scrollbars=yes,width=850,height=300')" class="btnType04">���߱�����Է�</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>