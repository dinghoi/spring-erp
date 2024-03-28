<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Response.CharSet = "EUC-KR"
Response.CodePage = "949"
Response.ContentType = "text/html;charset=euc-kr"
Response.CodePage = "949"

Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	run_month = Request("run_month")
	transit_id = Request("transit_id")
	view_c = Request("view_c")
	view_d = Request("view_d")
	use_man = Request("use_man")
  else
	run_month = Request.form("run_month")
	transit_id = Request.form("transit_id")
	view_c = Request.form("view_c")
	view_d = Request.form("view_d")
	use_man = Request.form("use_man")
end if

if view_d = "" then
    view_d = "run"
end If

If run_month = "" Then
	run_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	transit_id = "����"
    view_c = "total"
    view_d = "run"
	use_man = ""
End If

from_date = mid(run_month,1,4) + "-" + mid(run_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = run_month

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' �����Ǻ�
posi_sql = " and transit_cost.mg_ce_id = '" + user_id + "'"

'Response.write view_c
'Response.write cost_grade
'Response.write position

if position = "����" then
	view_condi = "����"
end If

'"��ȭ���� ����"�� ��� "��ȭ���� ����" ���絵 Ȯ�� �����ϰ� �߰�(������ �븮 ��û)[����ȣ_20210809]
if position = "��Ʈ��" then
	if view_c = "total" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (transit_cost.org_name = '��ȭ����ȣ��' or transit_cost.org_name = '��ȭ��������') "&chr(13)
		ElseIf org_name = "��ȭ���� ����" Then
			posi_sql = " and (transit_cost.org_name = '"&org_name&"' OR transit_cost.org_name = '��ȭ���� ����') "&chr(13)
		else
			posi_sql = " and transit_cost.org_name = '"&org_name&"' "&chr(13)
		end if
	else
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (transit_cost.org_name = '��ȭ����ȣ��' or transit_cost.org_name = '��ȭ��������') and memb.user_name like '%"&use_man&"%' "&chr(13)
		ElseIf org_name = "��ȭ���� ����" Then
			posi_sql = " and (transit_cost.org_name = '"&org_name&"' OR transit_cost.org_name = '��ȭ���� ����') and memb.user_name like '%"&use_man&"%' "&chr(13)
		else
			posi_sql = " and transit_cost.org_name = '"&org_name&"' and memb.user_name like '%"&use_man&"%' "&chr(13)
		end if
	end if
end if

if position = "����" then
	if view_c = "total" then
        posi_sql = " and transit_cost.team = '"&team&"' "&chr(13)
	else
        posi_sql = " and transit_cost.team =  '"&team&"' and memb.user_name like '%"&use_man&"%' "&chr(13)
	end if
end if

if position = "�������" or cost_grade = "2" then
    if view_c = "total" then
        'posi_sql = " and transit_cost.saupbu = '"&saupbu&"' "&chr(13)
        posi_sql = " and transit_cost.saupbu = emp_master.emp_saupbu "&chr(13)
    else
        'posi_sql = " and transit_cost.saupbu = '"&saupbu&"' and memb.user_name like '%"&use_man&"%' "&chr(13)
        posi_sql = " and transit_cost.saupbu = emp_master.emp_saupbu and memb.user_name like '%"&use_man&"%' "&chr(13)
    end if
end if

if position = "������" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and transit_cost.bonbu = '"&bonbu&"' "&chr(13)
 	else
		posi_sql = " and transit_cost.bonbu = '"&bonbu&"' and memb.user_name like '%"&use_man&"%' "&chr(13)
	end if
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "��ü"
  	if view_c = "total" then
		posi_sql = ""
 	else
		posi_sql = " and memb.user_name like '%"&use_man&"%'"
	end if
end if

if transit_id = "����" then
	transit_sql = " and (transit_cost.car_owner = '����' or transit_cost.car_owner = 'ȸ��') "&chr(13)
else
	transit_sql = " and (transit_cost.car_owner = '���߱���') "&chr(13)
end if

' ���Ǻ� ��ȸ.........
base_sql = "    select *                                          "&chr(13)&_
           "      from transit_cost                               "&chr(13)&_
           "inner join (SELECT user_id, user_name FROM memb) memb "&chr(13)&_
           "        on transit_cost.mg_ce_id = memb.user_id       "&chr(13)&_
           "inner join emp_master                                 "&chr(13)&_
           "        ON emp_master.emp_no = memb.user_id           "&chr(13)


if view_d = "run" then
    date_sql = " where (run_date >= '" + from_date  + "' and run_date <= '" + to_date  + "') "&chr(13)
    order_sql = " ORDER BY memb.user_name asc, run_date desc, run_seq desc "&chr(13)
end if
if view_d = "reg" then
    date_sql = " where (transit_cost.reg_date >= '" + from_date  + " 00:00:00' and transit_cost.reg_date <='" + to_date  + " 23:59:59') "&chr(13)
    order_sql = " ORDER BY memb.user_name asc, transit_cost.reg_date desc, run_seq desc "&chr(13)
end if

sql = "    select count(*)                                   "&chr(13)&_
      "      from transit_cost                               "&chr(13)&_
      "inner join (SELECT user_id, user_name FROM memb) memb "&chr(13)&_
      "        on transit_cost.mg_ce_id = memb.user_id       "&chr(13)&_
      "inner join emp_master                                 "&chr(13)&_
      "        ON emp_master.emp_no = memb.user_id           " + date_sql + posi_sql + transit_sql
Set RsCount = Dbconn.Execute (sql)
'Response.write Sql

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "       select sum(far) as far                            "&chr(13)&_
      "              ,sum(oil_price) as oil_price               "&chr(13)&_
      "              ,sum(fare) as fare                         "&chr(13)&_
      "              ,sum(repair_cost) as repair_cost           "&chr(13)&_
      "              ,sum(parking) as parking                   "&chr(13)&_
      "              ,sum(toll) as toll                         "&chr(13)&_
      "          from transit_cost                              "&chr(13)&_
      "   inner join (SELECT user_id, user_name FROM memb) memb "&chr(13)&_
      "           on transit_cost.mg_ce_id = memb.user_id       "&chr(13)&_
      "   inner join emp_master                                 "&chr(13)&_
      "           ON emp_master.emp_no = memb.user_id           "&chr(13)&_
      date_sql     &_
      posi_sql     &_
      transit_sql  &_
      " and cancel_yn = 'N'"

Set rs_sum = Dbconn.Execute (sql)
if rs_sum("far") = "" or isnull(rs_sum("far")) then
	sum_far         = 0
	sum_oil_price   = 0
	sum_fare        = 0
	sum_repair_cost = 0
	sum_parking     = 0
	sum_toll        = 0
  else
	sum_far         = rs_sum("far")
	sum_oil_price   = rs_sum("oil_price")
	sum_fare        = rs_sum("fare")
	sum_repair_cost = rs_sum("repair_cost")
	sum_parking     = rs_sum("parking")
	sum_toll        = rs_sum("toll")
end If

sql = base_sql + date_sql + posi_sql + transit_sql + order_sql + " limit "& stpage & "," &pgsize

'Response.write sql
'Response.end

Rs.Open Sql, Dbconn, 1

title_line = "����� ����"
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
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.run_month.value == "") {
					alert ("�������� �Է��ϼ���");
					return false;
				}
				return true;
			}
			function condi_view() {
                <% if position <> "����" or cost_grade = "0" then %>
                    if (eval("document.frm.view_c[0].checked")) {
                        document.getElementById('use_man_view').style.display = 'none';
                    }
                    if (eval("document.frm.view_c[1].checked")) {
                        document.getElementById('use_man_view').style.display = '';
                    }
                <% end if %>
			}
		</script>

	</head>
	<body onLoad="condi_view()">
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
                                    <input type="radio" name="view_d" value="run" <% if view_d = "run" then %>checked<% end if %> style="width:20px">
                                    <strong>������&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <% if view_d = "reg" then %>checked<% end if %> style="width:20px">
                                    <strong>�߱޳��&nbsp;</strong>

                                    : <input name="run_month" type="text" value="<%=run_month%>" style="width:70px">
                                    (��201401)
								</label>
								<label>
                              	<input type="radio" name="transit_id" value="����" <% if transit_id = "����" then %>checked<% end if %> style="width:20px">
                                ������������
                                <input type="radio" name="transit_id" value="����" <% if transit_id = "����" then %>checked<% end if %> style="width:20px">
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
                                <input type="radio" name="view_c" value="total" <%if view_c = "total" then %>checked <%end if %> style="width:20px" onClick="condi_view();">
                                    ������ü
                                <input type="radio" name="view_c" value="reg_id" <%if view_c = "reg_id" then %>checked <%end if %> style="width:20px" onClick="condi_view();">
                                    ���κ�
								<%End If%>
                                </label>
								<label>
                                	<input name="use_man" type="text" value="<%=use_man%>" style="width:70px; display:none" id="use_man_view">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
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
						do until rs.eof
							if rs("car_owner") = "���߱���" then
								car_gubun = rs("transit")
							  else
								car_gubun = rs("car_owner") + "<br>" + rs("oil_kind")
							end if
							run_km = rs("far")

							if rs("cancel_yn") = "Y" then
								cancel_view = "���"
							  else
							  	cancel_view = "����"
							end if
							if rs("start_km") = "" or isnull(rs("start_km")) then
								start_view = 0
							  else
							  	start_view = rs("start_km")
							end if
							if rs("end_km") = "" or isnull(rs("end_km")) then
								end_view = 0
							  else
							  	end_view = rs("end_km")
							end if
						    %>
							<%
                            ' 5�� ���� ���� �Է°� ����...
                            chk_slip_month = mid(rs("run_date"),1,7)
                            chk_reg_month = mid(rs("reg_date"),1,7)
                            chk_reg_day = mid(rs("reg_date"),9,2)

                            if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                bgcolor = "burlywood"
                            else
                                bgcolor = "#f8f8f8"
                            end if
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
								<td class="first"><%=rs("user_name")%></td>
                                <td><%=rs("run_date")%></td>
                                <td><%=mid(rs("reg_date"),1,10)%></td>
								<td><%=car_gubun%></td>
								<td><%=rs("start_company")%>-<%=rs("start_point")%></td>
								<td><%=rs("end_company")%>-<%=rs("end_point")%></td>
								<td><%=rs("run_memo")%>&nbsp;</td>
								<td class="right"><%=formatnumber(start_view,0)%></td>
								<td class="right"><%=formatnumber(end_view,0)%></td>
								<td class="right"><%=formatnumber(run_km,0)%></td>
								<td class="right"><%=formatnumber(rs("repair_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("fare"),0)%></td>
								<td class="right"><%=formatnumber(rs("oil_price"),0)%></td>
								<td class="right"><%=formatnumber(rs("parking"),0)%></td>
								<td class="right"><%=formatnumber(rs("toll"),0)%></td>
								<td><%=cancel_view%></td>
								<td>
                                    <% if rs("end_yn") <> "Y" then	%>
                                        <% if rs("car_owner") = "���߱���" then  %>
                                            <% if rs("mg_ce_id") = user_id then	%>
                                                <a href="#" onClick="pop_Window('mass_transit_add.asp?mg_ce_id=<%=rs("mg_ce_id")%>&run_date=<%=rs("run_date")%>&run_seq=<%=rs("run_seq")%>&u_type=<%="U"%>','mass_transit_add_pop','scrollbars=yes,width=850,height=350')">����</a>
                                            <% else	%>
                                                <a href="#" onClick="pop_Window('mass_transit_cancel.asp?mg_ce_id=<%=rs("mg_ce_id")%>&run_date=<%=rs("run_date")%>&run_seq=<%=rs("run_seq")%>&u_type=<%="U"%>','mass_transit_cancel_pop','scrollbars=yes,width=850,height=350')">����</a>
                                            <% end if	%>
                                        <% else  %>
                                            <% if rs("mg_ce_id") = user_id then	%>
                                                <a href="#" onClick="pop_Window('car_drive_add.asp?mg_ce_id=<%=rs("mg_ce_id")%>&run_date=<%=rs("run_date")%>&run_seq=<%=rs("run_seq")%>&u_type=<%="U"%>','car_drive_add_pop','scrollbars=yes,width=900,height=500')">����</a>
                                            <% else	%>
                                                <a href="#" onClick="pop_Window('car_drive_cancel.asp?mg_ce_id=<%=rs("mg_ce_id")%>&run_date=<%=rs("run_date")%>&run_seq=<%=rs("run_seq")%>&u_type=<%="U"%>','car_drive_cancel_pop','scrollbars=yes,width=900,height=470')">����</a>
                                            <% end if %>
                                        <% end if %>
                                    <% else	%>
                                        ����
                                    <% end if 	%>
                                </td>
							</tr>
						    <%
							rs.movenext()
						loop
						rs.close()

						if tottal_record <> 0 then
						%>
							<tr>
								<th class="first">��</th>
								<th colspan="3"><%=tottal_record%>&nbsp;��</th>
								<th colspan="13">����Ÿ� :&nbsp;<%=formatnumber(sum_far,0)%>&nbsp;KM&nbsp;&nbsp;,&nbsp;&nbsp;������ :&nbsp;<%=formatnumber(sum_repair_cost,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;���߱���� :&nbsp;<%=formatnumber(sum_fare,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;�����ݾ� :&nbsp;<%=formatnumber(sum_oil_price,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;������ :&nbsp;<%=formatnumber(sum_parking,0)%>&nbsp;&nbsp;,&nbsp;&nbsp;����� :&nbsp;<%=formatnumber(sum_toll,0)%></th>
							</tr>
						<%
						rs_sum.close()
						end if
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
                    <div class="btnCenter">
                        <a href="transit_cost_excel.asp?run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>" class="btnType04">�����ٿ�ε�</a>
                    </div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="transit_cost_mg.asp?page=<%=first_page%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[ó��]</a>
                        <% if intstart > 1 then %>
                            <a href="transit_cost_mg.asp?page=<%=intstart -1%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[����]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="transit_cost_mg.asp?page=<%=i%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if 	intend < total_page then %>
                            <a href="transit_cost_mg.asp?page=<%=intend+1%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[����]</a>
                            <a href="transit_cost_mg.asp?page=<%=total_page%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[������]</a>
                            <%	else %>
                            [����]&nbsp;[������]
                        <% end if %>
                    </div>
                    </td>
				    <td width="25%">
                    <div class="btnCenter">
                        <a href="#" onClick="pop_Window('car_drive_add.asp','car_drive_add_popup','scrollbars=yes,width=900,height=450')" class="btnType04">�������������Է�</a>
                        <a href="#" onClick="pop_Window('mass_transit_add.asp','mass_transit_add_popup','scrollbars=yes,width=850,height=300')" class="btnType04">���߱�����Է�</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>
