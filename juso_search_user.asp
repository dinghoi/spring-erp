<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
check_ok = "y"
check_no = "n"
'dim ddd_tab(30)

company = Request("company")
if c_grade = "5" then
	company = user_name
end if
dept = Request.form("dept")
tel_ddd = Request.form("tel_ddd")
tel_no1 = Request.form("tel_no1")
tel_no2 = Request.form("tel_no2")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_ddd = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if dept = "" or isnull(dept) then
	first_view = "N"
	sql = "select * from juso_list where company = '" + company + "' and dept = '" + dept + "' ORDER BY dept ASC"
  else
	first_view = "Y"
	sql = "select * from juso_list where company = '" + company + "' and dept like '%" + dept + "%' ORDER BY dept ASC"
end if
rs.open sql, Dbconn, 1

title_line = "�ּҷ� DB �˻�"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�ּ�DB �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
		function jusocode(tel_ddd,tel_no1,tel_no2,company,dept,sido,gugun,dong,addr,mg_ce_id,mg_ce,team,reside_place,reside_company,view_ok){
				opener.document.frm.tel_ddd.value = tel_ddd;
				opener.document.frm.tel_no1.value = tel_no1;
				opener.document.frm.tel_no2.value = tel_no2;
				opener.document.frm.company.value = company;
				opener.document.frm.dept.value = dept;
				opener.document.frm.sido.value = sido;
				opener.document.frm.gugun.value = gugun;
				opener.document.frm.dong.value = dong;
				opener.document.frm.addr.value = addr;
				opener.document.frm.mg_ce_id.value = mg_ce_id;
				opener.document.frm.mg_ce.value = mg_ce;
				opener.document.frm.team.value = team;
				opener.document.frm.reside_place.value = reside_place;
//				opener.document.frm.reside_company.value = reside_company;
//				if (view_ok=="y"){
//					opener.document.frm.area_view.style.display="none" ;}
//				else {
//					opener.document.frm.area_view.style.display="" ;}
				opener.document.frm.acpt_user.focus();
				window.close();
			
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.dept.value =="") {
					alert('�������� �Է��ϼ���');
					frm.dept.focus();
					return false;}
				if(document.frm.tel_ddd.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_ddd.focus();
					return false;}
				if(document.frm.tel_no1.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no1.focus();
					return false;}
				if(document.frm.tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no2.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="juso_search_user.asp?company=<%=company%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>ȸ�� : </strong>
        					<%=company%>
							<strong>������ : </strong>	
                                <input name="dept" type="text" id="dept" value="<%=dept%>" size="20" onKeyUp="checklength(this,50)" style="text-align:left; ime-mode:active">
							<strong>&nbsp;��ȭ��ȣ : </strong>
							<% 
                                Sql="select * from etc_code where etc_type = '71' and used_sw = 'Y' order by etc_name asc"
                                Rs_ddd.Open Sql, Dbconn, 1
                            %>
                            	<select name="tel_ddd" id="tel_ddd" style="width:50px">
                            <% 
                                do until rs_ddd.eof 
                            %>
                              		<option value='<%=rs_ddd("etc_name")%>' <%If rs_ddd("etc_name") = tel_ddd then %>selected<% end if %>><%=rs_ddd("etc_name")%></option>
                            <%
                                    rs_ddd.movenext()
                                loop
                                rs_ddd.close()						
                            %>
                            	</select>)
                          		<input name="tel_no1" type="text" id="tel_no1" value="<%=tel_no1%>" size="4" maxlength="4" onKeyPress="onlynum();">-
                          		<input name="tel_no2" type="text" id="tel_no2" value="<%=tel_no2%>" size="4" maxlength="4" onKeyPress="onlynum();">
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��ȭ��ȣ</th>
								<th scope="col">ȸ��</th>
								<th scope="col">������</th>
								<th scope="col">�ּ�</th>
								<th scope="col">�����</th>
							</tr>
						</thead>
						<tbody>
						<%
						if first_view = "Y" then
						%>
							<tr>
								<td class="first"><%=tel_ddd%>&nbsp;<%=tel_no1%>&nbsp;<%=tel_no2%></td>
								<td><%=company%></td>
								<td>
                                <a href="#" onClick="jusocode('<%=tel_ddd%>','<%=tel_no1%>','<%=tel_no2%>','<%=company%>','<%=dept%>','','','','','','','<%=check_no%>');"><%=dept%></a>
                                </td>
								<td class="left">���� �ּҷ� �����ϰ� �Է��� ���� ���</td>
								<td>&nbsp;</td>
							</tr>
          					<% 
							do until rs.eof or rs.bof

								Sql_memb="select * from memb where user_id = '"&rs("mg_ce_id")&"'"
								Rs_memb.Open Sql_memb, Dbconn, 1
								mg_ce_id = rs("mg_ce_id")
								if 	rs_memb.eof or rs_memb.bof then
									mg_ce_id = ""
									mg_ce = "�����"
									reside_place = ""
									reside_company = ""
									reside_sw = "0"
									team = ""
								  else
									mg_ce = rs_memb("user_name")
									reside_place = rs_memb("reside_place")
									reside_company = rs_memb("reside_company")
									reside_sw = rs_memb("reside")
									team = rs_memb("team")			
								end if
								rs_memb.close()
							%>
							<tr>
								<td class="first"><%=rs("tel_ddd")%>&nbsp;<%=rs("tel_no1")%>&nbsp;<%=rs("tel_no2")%></td>
								<td><%=rs("company")%></td>
								<td>
                                <a href="#" onClick="jusocode('<%=tel_ddd%>','<%=tel_no1%>','<%=tel_no2%>','<%=rs("company")%>','<%=rs("dept")%>','<%=rs("sido")%>','<%=rs("gugun")%>','<%=rs("dong")%>','<%=rs("addr")%>','<%=mg_ce_id%>','<%=mg_ce%>','<%=team%>','<%=reside_place%>','<%=reside_company%>','<%=check_ok%>');"><%=rs("dept")%></a>
                                </td>
								<td class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></td>
								<td><%=mg_ce%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						end if
						%>
						</tbody>
					</table>
				</div>
			</div>				
	</div>        				
	</form>
	</body>
</html>

