<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

slip_month=Request.form("slip_month")
view_c = Request.form("view_c")
view_d = Request.form("view_d")
emp_name=Request.form("emp_name")

if view_d = "" then
    view_d = "slip"
end if

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
    view_c = "total"
    view_d = "slip"
	emp_name = ""
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = slip_month

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' �����Ǻ�
posi_sql = " and general_cost.emp_no = '" + user_id + "'"

if position = "����" then
	view_condi = "����"
end if

if position = "��Ʈ��" then
	if view_c = "total" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (general_cost.org_name = '��ȭ����ȣ��' or general_cost.org_name = '��ȭ��������') "
		  else
			posi_sql = " and general_cost.org_name = '"&org_name&"'"
		end if
	else
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (general_cost.org_name = '��ȭ����ȣ��' or general_cost.org_name = '��ȭ��������') and general_cost.emp_name like '%"&emp_name&"%'"
		  else
			posi_sql = " and general_cost.org_name = '"&org_name&"' and general_cost.emp_name like '%"&emp_name&"%'"
		end if
	end if
end if

if position = "����" then
	if view_c = "total" then
        'posi_sql = " and team = '"&team&"'"
        posi_sql = " and (team = '"&team&"' or reside_place = '"&team&"') "&chr(13)
    else
        'posi_sql = " and team = '"&team&"' and general_cost.emp_name like '%"&emp_name&"%'"
        posi_sql = " and (team = '"&team&"' or reside_place = '"&team&"') and general_cost.emp_name like '%"&emp_name&"%' "&chr(13)
	end if
end if

if position = "�������" or cost_grade = "2" then
	if view_c = "total" then
        'posi_sql = " and saupbu = '"&saupbu&"' "&chr(13)
        posi_sql = " and saupbu = emp_master.emp_saupbu "&chr(13)
	else
        'posi_sql = " and saupbu = '"&saupbu&"' and emp_name like '%"&emp_name&"%' "&chr(13)
        posi_sql = " and saupbu = emp_master.emp_saupbu and general_cost.emp_name like '%" & emp_name & "%' "&chr(13)
	end if
end if

if position = "������" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and general_cost.bonbu = '"&bonbu&"'"&chr(13)
 	else
		posi_sql = " and general_cost.bonbu = '"&bonbu&"' and general_cost.emp_name like '%"&emp_name&"%'"&chr(13)
	end if
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "��ü"
  	if view_c = "total" then
		posi_sql = " "
 	  else
		posi_sql = " and general_cost.emp_name like '%"&emp_name&"%'"
	end if
end if

' ���Ǻ� ��ȸ.........
base_sql = "     select *                                           "&chr(13)&_
           "       from general_cost                                "&chr(13)&_
           " inner join emp_master                                  "&chr(13)&_
           "         ON emp_master.emp_no =  general_cost.emp_no    "&chr(13)&_
		   " inner join emp_org_mst                                  "&chr(13)&_
           "         ON emp_master.emp_org_code = emp_org_mst.org_code  "&chr(13)&_
           "      where (cost_reg = '0')                            "&chr(13)&_
           "        and (tax_bill_yn <> 'Y' or isnull(tax_bill_yn)) "&chr(13)&_
           "        and (slip_gubun = '���')                       "&chr(13)

if view_d = "slip" then
    base_sql = base_sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "&chr(13)
    order_sql = " ORDER BY general_cost.org_name, general_cost.emp_name, general_cost.slip_date ASC"
end if
if view_d = "reg" then
    base_sql = base_sql & " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59') "&chr(13)
    order_sql = " ORDER BY general_cost.org_name, general_cost.emp_name, general_cost.reg_date ASC"
end If

sql = base_sql & posi_sql & order_sql

Rs.Open Sql, Dbconn, 1

title_line = "�Ϲݰ�� ����"
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
				if (document.frm.slip_month.value == "") {
					alert ("�߻������ �Է��ϼ���.");
					return false;
				}
				return true;
			}
			function condi_view() {
                <% if position <> "����" or cost_grade = "0" then %>
                    if (eval("document.frm.view_c[0].checked")) {
                        document.getElementById('emp_name_view').style.display = 'none';
                    }
                    if (eval("document.frm.view_c[1].checked")) {
                        document.getElementById('emp_name_view').style.display = '';
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
				<form action="/general_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
                                    <input type="radio" name="view_d" value="slip" <% if view_d = "slip" then %>checked<% end if %> style="width:25px">
                                    <strong>�߻����&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <% if view_d = "reg" then %>checked<% end if %> style="width:25px">
                                    <strong>�߱޳��&nbsp;</strong>

                                    : <input name="slip_month" type="text" value="<%=slip_month%>" style="width:70px">
                                    (��201401)
								</label>
								<label>
								    <strong>��ȸ���� : </strong><%=view_grade%>
								</label>
								<label>
								<strong>��ȸ���� : </strong>
                                <% if position = "����" and cost_grade <> "0" then %>
                                    <%=view_condi%>
                                <% else	%>
                                    <input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                    ������ü
                                    <input type="radio" name="view_c" value="reg_id" <% if view_c = "reg_id" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                    ���κ�
                                <% end if %>
								</label>
								<label>
                                	<input name="emp_name" type="text" value="<%=emp_name%>" style="width:70px; display:none" id="emp_name_view">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="11%" >
							<col width="5%" >
							<col width="7%" >
							<col width="*" >
							<col width="5%" >
							<col width="4%" >
							<col width="3%" >
							<col width="16%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ҽ�</th>
								<th scope="col">�����</th>
								<th scope="col">�߻�����</th>
								<th scope="col">�߱�����</th>
								<th scope="col">��뱸��</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">����NO</th>
								<th scope="col">��û�ݾ�</th>
								<th scope="col">���ó</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">���</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
						cost_sum = 0
						pay_sum = 0
						mi_pay_sum = 0
						cancel_sum = 0
						do until rs.eof
							cost_sum = cost_sum + rs("cost")
							if rs("cancel_yn") = "Y" then
								cancel_sum = cancel_sum + rs("cost")
							end if
							if rs("cancel_yn") <> "Y" then
								if rs("pay_yn") = "Y" then
									pay_sum = pay_sum + rs("cost")
								  else
									mi_pay_sum = mi_pay_sum + rs("cost")
								end if
							end if

							if rs("pay_yn") = "Y" then
								pay_yn = "����"
							  else
							  	pay_yn = "������"
							end if
							if rs("cancel_yn") = "Y" then
								cancel_yn = "���"
							  else
							  	cancel_yn = "����"
							end if
                            %>
                            <%
                            ' 5�� ���� ���� �Է°� ����...
                            chk_slip_month = mid(rs("slip_date"),1,7)
                            chk_reg_month = mid(rs("reg_date"),1,7)
                            chk_reg_day = mid(rs("reg_date"),9,2)

                            if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                bgcolor = "burlywood"
                            else
                                bgcolor = "#f8f8f8"
                            end if
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
                                <td class="first"><%=rs("org_name")%></td>
                                <td><%=rs("emp_name")%>&nbsp;<%=rs("emp_grade")%></td>
                                <td><%=rs("slip_date")%></td>
                                <td><%=mid(rs("reg_date"),1,10)%></td>
                                <td><%=rs("account")%></td>
                                <td><%=rs("company")%></td>
                                <td><%=rs("sign_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cost"),0)%></td>
                                <td><%=rs("customer")%></td>
                                <td><%=pay_yn%></td>
                                <td><%=cancel_yn%></td>
                                <td><%=rs("pl_yn")%></td>
                                <td><%=rs("slip_memo")%></td>
                                <td>
                                <% if rs("end_yn") <> "Y" then %>
                                    <% if rs("emp_no") = user_id or cost_grade = "0" then	%>
                                        <% if cost_grade = "5" or cost_grade = "6" then	%>
                                            <a href="#" onClick="pop_Window('/general_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','general_cost_add_pop','scrollbars=yes,width=900,height=350')">����</a>
                                        <% else %>
                                            <a href="#" onClick="pop_Window('/common_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','common_cost_add_pop','scrollbars=yes,width=900,height=360')">����</a>
                                        <% end if	%>
                                    <% else	%>
                                        <% if cost_grade = "5" or cost_grade = "6" then	%>
                                            <a href="#" onClick="pop_Window('/general_cost_cancel.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','general_cost_cancel_pop','scrollbars=yes,width=900,height=350')">����</a>
                                        <% else	%>
                                            <a href="#" onClick="pop_Window('/common_cost_cancel.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','general_cost_cancel_pop','scrollbars=yes,width=900,height=360')">����</a>
                                        <% end if	%>
                                    <% end if	%>
                                <% else	%>
                                    ����
                                <% end if %>
                                </td>
                            </tr>
                            <%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th colspan="1" class="first">�� ��</th>
								<th colspan="4">��û�ݾ� :&nbsp;<%=formatnumber(cost_sum,0)%></th>
								<th colspan="3">�����޾� :&nbsp;<%=formatnumber(mi_pay_sum,0)%></th>
							  	<th colspan="3">�����޾� :&nbsp;<%=formatnumber(pay_sum,0)%></th>
							  	<th colspan="3">��ұݾ� :&nbsp;<%=formatnumber(cancel_sum,0)%></th>
						  	</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnCenter" style="text-align:left;">
                        <a href="/general_cost_excel.asp?slip_month=<%=slip_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>" class="btnType04">�����ٿ�ε�</a>
					<%If cost_grade = "0" Then%>
						<a href="/cost/general_cost_excel_upload.asp" class="btnType04">�ϰ� ���ε�</a>
					<%End If%>
					</div>
					</td>
					<td width="40%">
                    </td>
				    <td width="25%">
					<div class="btnCenter">
                        <% if cost_grade = "5" or cost_grade = "6" or user_id="101227" then '���뼮 19.01.11 �䱸 %>
                            <a href="#" onClick="pop_Window('/general_cost_add.asp','general_cost_add_pop','scrollbars=yes,width=900,height=300')" class="btnType04">CE ����Է�</a>
                        <% elseif cost_grade = "0" then	%>
                            <a href="#" onClick="pop_Window('/general_cost_add.asp','general_cost_add_pop','scrollbars=yes,width=900,height=300')" class="btnType04">CE ����Է�</a>
                            <a href="#" onClick="pop_Window('/common_cost_add.asp','common_cost_add_pop','scrollbars=yes,width=900,height=300')" class="btnType04">�����װ����� ����Է�</a>
                        <% else %>
                            <a href="#" onClick="pop_Window('/common_cost_add.asp','common_cost_add_pop','scrollbars=yes,width=900,height=300')" class="btnType04">�����װ����� ����Է�</a>
                        <% end if %>
					</div>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>
