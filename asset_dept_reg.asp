<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim company_tab(50,2)
company = request("company")
u_type = request("u_type")

if asset_company <> "00" then
	company = asset_company
end if

high_org = ""
org_first = ""
org_second = ""
dept_name = ""
person = ""
sido = ""
gugun = ""
dong = ""
addr = ""
tel_ddd = ""
tel_no1 = ""
tel_no2 = ""
internet_no = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "�����ڵ� ���"
if u_type = "U" then

	dept_code = request("dept_code")

	sql = "select * from asset_dept where company = '" + company + "' and dept_code = '" + dept_code + "'"
	set rs = dbconn.execute(sql)
	
	high_org = rs("high_org")
	org_first = rs("org_first")
	org_second = rs("org_second")
	dept_name = rs("dept_name")
	person = rs("person")
	sido = rs("sido")
	gugun = rs("gugun")
	dong = rs("dong")
	addr = rs("addr")
	tel_ddd = rs("tel_ddd")
	tel_no1 = rs("tel_no1")
	tel_no2 = rs("tel_no2")
	internet_no = rs("internet_no")

	rs.close()

	title_line = "�ڻ��ڵ� ����"
end if

title_line = "�����ڵ� ���"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=bill_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=inout_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.company.value =="") {
					alert('ȸ�縦 �Է��ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.high_org.value =="") {
					alert('���������� �����ϼ���');
					frm.high_org.focus();
					return false;}
				if(document.frm.org_first.value =="") {
					alert('���θ��� �����ϼ���');
					frm.org_first.focus();
					return false;}
//				if(document.frm.org_second.value =="") {
//					alert('������� �Է��ϼ���');
//					frm.org_second.focus();
//					return false;}
//				if(document.frm.dept_name.value =="") {
//					alert('�������� �Է��ϼ���');
//					frm.dept_name.focus();
//					return false;}
				if(document.frm.person.value =="") {
					alert('����ڸ� �Է��ϼ���');
					frm.person.focus();
					return false;}
				if(document.frm.sido.value =="") {
					alert('������ȸ�� �ϼ���');
//					frm.area_view.focus();
					return false;}
				if(document.frm.addr.value =="") {
					alert('������ �Է��ϼ���');
					frm.addr.focus();
					return false;}
				if(document.frm.tel_ddd.value =="") {
					alert('DDD�� �Է��ϼ���');
					frm.tel_ddd.focus();
					return false;}
				if(document.frm.tel_no1.value =="") {
					alert('��ȭ���� �Է��ϼ���');
					frm.tel_no1.focus();
					return false;}
				if(document.frm.tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no2.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body onload="specview()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="asset_dept_reg_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�Ҽ�ȸ��</th>
								<td class="left">
								  <%
                                    if	asset_company = "00" then
                                        k = 0
                                        Sql="select * from etc_code where etc_type = '75' and used_sw = 'Y' order by etc_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                        while not rs_etc.eof
                                            k = k + 1
                                            company_tab(k,1) = rs_etc("etc_name")
                                            company_tab(k,2) = mid(rs_etc("etc_code"),3,2)
                                            rs_etc.movenext()
                                        Wend
                                        rs_etc.close()						
                                    %>
                                  <select name="company" id="company" style="width:150px">
                                    <option value="">����</option>
                                    <% 
                                            for kk = 1 to k
                                        %>
                                    <option value='<%=company_tab(kk,2)%>' <%If company_tab(kk,2) = asset_company then %>selected<% end if %>><%=company_tab(kk,1)%></option>
                                    <%
                                            next
                                        %>
                                  </select>
                                <%		else	%>
                                    <%=user_name%>
                                    <input name="company" type="hidden" id="company" value="<%=company%>">
                                <%	end if	%>
                                </td>
								<th>��������</th>
								<td class="left">
								  <%
                                        Sql="select * from org_code where org_company = '" + company + "' and org_gubun = '1' order by org_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="high_org" id="select2" style="width:150px">
                                    <option value="">����</option>
                                    <% 
                                        While not rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("org_name")%>' <%If rs_etc("org_name") = high_org then %>selected<% end if %>><%=rs_etc("org_name")%></option>
                                    <%
                                            rs_etc.movenext()  
                                        Wend 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                    <input name="dept_code" type="hidden" id="dept_code" value="<%=dept_code%>">
                                </td>
							</tr>
							<tr>
								<th class="first">���θ�</th>
								<td class="left">
								  <%
                                        Sql="select * from org_code where org_company = '" + company + "' and org_gubun = '2' order by org_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="org_first" id="org_first" style="width:150px">
                                    <option value="">����</option>
                                    <% 
                                        While not rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("org_name")%>' <%If rs_etc("org_name") = org_first then %>selected<% end if %>><%=rs_etc("org_name")%></option>
                                    <%
                                            rs_etc.movenext()  
                                        Wend 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                </td>
								<th>�����</th>
								<td class="left"><input name="org_second" type="text" id="org_second" style="width:150px" onKeyUp="checklength(this,30)" value="<%=org_second%>"></td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left"><input name="dept_name" type="text" id="dept_name" style="width:150px" onKeyUp="checklength(this,30)" value="<%=dept_name%>"></td>
								<th>�����</th>
								<td class="left"><input name="person" type="text" id="person" style="width:150px" onKeyUp="checklength(this,20)" value="<%=person%>"></td>
							</tr>
							<tr>
								<th class="first">�ּ�</th>
								<td  colspan="3" class="left">
                                  <input name="sido" type="text" id="sido" size="6" maxlength="6" readonly="true" value="<%=sido%>">&nbsp;
                                  <input name="gugun" type="text" id="gugun" size="20" maxlength="20" readonly="true" value="<%=gugun%>">&nbsp;
                                  <input name="dong" type="text" id="dong" size="20" maxlength="20" readonly="true" value="<%=dong%>">&nbsp;
								  <a href="#" class="btnType03" onClick="pop_Window('area_search.asp','areacode','scrollbars=yes,width=600,height=400')">������ȸ</a>
                                  <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="">
                                  <input name="mg_ce" type="hidden" id="mg_ce" value="">
                                  <input name="team" type="hidden" id="team" value="">
                                  <input name="reside_place" type="hidden" id="reside_place" value="">
                                </td>
							</tr>
							<tr>
								<th class="first">����</th>
								<td  colspan="3" class="left"><input name="addr" type="text" id="addr" style="width:400px" onKeyUp="checklength(this,50)" value="<%=addr%>"></td>
							</tr>
							<tr>
								<th class="first">��ȭ��ȣ</th>
								<td class="left">
                                <input name="tel_ddd" type="text" id="tel_ddd2" size="3" maxlength="3" value="<%=tel_ddd%>">-
                                <input name="tel_no1" type="text" id="tel_no12" size="4" maxlength="4" value="<%=tel_no1%>">-
                                <input name="tel_no2" type="text" id="tel_no22" size="4" maxlength="4" value="<%=tel_no2%>">
                                </td>
								<th>���ͳ�NO</th>
								<td class="left"><input name="internet_no" type="text" id="internet_no" style="width:150px" onKeyUp="checklength(this,50)" value="<%=internet_no%>"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

