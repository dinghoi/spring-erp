<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

u_type = request("u_type")
user_id = request("user_id")

sql = "select * from memb where user_grade = 'ȸ��' and (asset_company = '' or isnull(asset_company)) order by user_name desc"
Rs.Open Sql, Dbconn, 1
'Response.write sql

if u_type = "U" then
	sql = "select * from memb where user_id = '" + user_id + "'"
	Set rs_etc=DbConn.Execute(Sql)
	user_id = rs_etc("user_id")
	user_name = rs_etc("user_name")
	pass = rs_etc("pass")
	reside = rs_etc("reside")
	grade = rs_etc("grade")
  else
	user_id = ""
	user_name = ""
	pass = ""
	reside = ""
	grade = ""
end if	

title_line = "ȸ�� ����� ��� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	 	<script src="/java/jquery-1.9.1.js"></script>
	 	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			function frmsubmit () {
				document.condi_frm.submit ();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.u_type.value != "U") {
					if(document.frm.user_id.value =="") {
						alert('���̵� �Է��ϼ���');
						frm.user_id.focus();
						return false;}}
				if(document.frm.company.value =="") {
					alert('ȸ����� �Է��ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.pass.value =="") {
					alert('��й�ȣ�� �Է��ϼ���');
					frm.pass.focus();
					return false;}

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.reside[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("��ȸ������ �����ϼ���");
					return false;
				}	

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.grade[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("������� �����ϼ���");
					return false;
				}	
			
				a=confirm('����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
			function frmcancel() 
				{
					document.frm.action = "com_user_mg.asp?u_type=''";
					document.frm.submit();
				}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="69%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="13%" >
				          <col width="*" >
				          <col width="12%" >
				          <col width="10%" >
				          <col width="10%" >
				          <col width="20%" >
				          <col width="10%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">���̵�</th>
				            <th scope="col">ȸ���</th>
				            <th scope="col">��й�ȣ</th>
				            <th scope="col">��ȸ����</th>
				            <th scope="col">�α���Ƚ��</th>
				            <th scope="col">�����α���</th>
				            <th scope="col">�������</th>
			              </tr>
			            </thead>
			            <tbody>
									<%
									do until rs.eof
										if rs("reside") = "9" then
											group_view = "�׷���ȸ"
											else
												group_view = "�ܵ���ȸ"
										end if
										if rs("grade") = "5" then
											use_view = "���"
											else
												use_view = "�̻��"
										end if
											%>
											<tr>
												<td class="first"><%=rs("user_id")%></td>
												<td><a href="com_user_mg.asp?user_id=<%=rs("user_id")%>&u_type=<%="U"%>"><%=rs("user_name")%></a></td>
												<td><%=rs("pass")%></td>
												<td>
												<% if rs("reside") = "9" then	%>
													<a href="#" onClick="pop_Window('view_group.asp?company=<%=rs("user_name")%>','view_group_pop','scrollbars=yes,width=400,height=500')"><%=group_view%></a>
												<%   else	%>
													<%=group_view%>
												<%  end if	%>
												</td>
												<td><%=formatnumber(rs("login_cnt"),0)%></td>
												<td><%=rs("login_date")%>&nbsp;</td>
												<td><%=use_view%></td>
											</tr>
											<%
										rs.movenext()
									loop
									%>
			            </tbody>
			          </table>
              </td>
				      <td width="1%" valign="top">&nbsp;</td>
				      <td width="30%" valign="top"><form method="post" name="frm" action="com_user_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="30%">���̵�</th>
				              <td class="left">
              				<% if u_type = "U" then	%>
                        <%=user_id%>
                      <% else	%>
                      	<input name="user_id" type="text" id="user_id" onKeyUp="checklength(this,20)" value="<%=user_id%>" style="width:130px">
                      <% end if	%>
                      </td>
			                </tr>
				            <tr>
				              <th>ȸ���</th>
				              <td class="left">
                      	<input name="company" type="text" id="company" style="width:130px" value="<%=user_name%>" readonly="true">
                        <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">ȸ����ȸ</a>
                      </td>
			                </tr>
				            <tr>
				              <th>��й�ȣ</th>
				              <td class="left"><input name="pass" type="text" id="pass" onKeyUp="checklength(this,20)" value="<%=pass%>" style="width:130px"></td>
			                </tr>
				            <tr>
				              <th>��ȸ����</th>
				              <td class="left">
                        <input type="radio" name="reside" value="0" <% if reside = "0" then %>checked<% end if %> style="width:40px" ID="Radio5">�ܵ���ȸ
  							  			<input type="radio" name="reside" value="9" <% if reside = "9" then %>checked<% end if %> style="width:40px" ID="Radio6">�׷���ȸ
                      </td>
			                </tr>
				            <tr>
				              <th>�������</th>
				              <td class="left">
                        <input type="radio" name="grade" value="5" <% if grade = "5" then %>checked<% end if %> style="width:40px" ID="Radio5">���
  							  			<input type="radio" name="grade" value="6" <% if grade = "6" then %>checked<% end if %> style="width:40px" ID="Radio6">�̻��
                      </td>
			                </tr>
			              </tbody>
			            </table>
							<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="old_user_id" value="<%=user_id%>" ID="Hidden1">
				        <div align=center>
                  <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                  <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcancel();" ID="Button1" NAME="Button1"></span>
                </div>
			          </form></td>
			        </tr>
				    	<tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
          </div>
			</div>				
	</div>        				
	</body>
</html>

