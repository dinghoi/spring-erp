<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_family.asp"

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	inc_yyyy = request.form("inc_yyyy")
  else
	inc_yyyy = request("inc_yyyy")
end if

if view_condi = "" then
	'inc_yyyy = mid(cstr(now()),1,4)
	inc_yyyy = cint(mid(now(),1,4)) - 1
	ck_sw = "n"
end if

' �ֱ�3���⵵ ���̺�� ����
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "��"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "��"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "��"


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_ann = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_dona = Server.CreateObject("ADODB.Recordset")
Set rs_duct = Server.CreateObject("ADODB.Recordset")
Set rs_cred = Server.CreateObject("ADODB.Recordset")
Set rs_hous = Server.CreateObject("ADODB.Recordset")
Set rs_houm = Server.CreateObject("ADODB.Recordset")
Set rs_savi = Server.CreateObject("ADODB.Recordset")
Set rs_other = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")

sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if not rs_year.eof then
       y_final =  rs_year("y_final") 
   else	   
	   y_final =  ""
end if
rs_year.close()	

sql = "select * from emp_family where family_empno = '"&emp_no&"' ORDER BY family_empno,family_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "�������� - ������ �ȳ� �� ���ǻ��� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_family.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <label>
                             <strong>��� : </strong>
                                <input name="emp_no" type="text" value="<%=emp_no%>" style="width:50px" readonly="true">
                                -
                                <input name="emp_name" type="text" value="<%=emp_name%>" style="width:60px" readonly="true">
                                </label>
                                <label>
                             <strong>���� : </strong>
                                <input name="emp_grade" type="text" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                -
                                <input name="emp_position" type="text" value="<%=emp_position%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>�Ի��� : </strong>
                                <input name="emp_in_date" type="text" value="<%=emp_in_date%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>�Ҽ� : </strong>
                                <input name="emp_company" type="text" value="<%=emp_company%>" style="width:90px" readonly="true">
                                -
                                <input name="emp_org_name" type="text" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                </label>
                             <strong>�ͼӳ⵵ : </strong>
                                <select name="inc_yyyy" id="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:70px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If inc_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="8%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
                            
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="8%" >
                            
                            <col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" scope="col" class="first">����</th>
                                <th rowspan="2" scope="col">����</th>
								<th rowspan="2" scope="col">�ֹε�Ϲ�ȣ</th>
								<th rowspan="2" scope="col">���ܱ���<br>����</th>
                                <th rowspan="2" scope="col">�ξ翩��</th>
								<th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
                                <th rowspan="2" scope="col">��Ÿ</th>
                                <th rowspan="2" scope="col">����</th>
							</tr>
                            <tr>
				                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">�����</th>
				                <th scope="col">����������</th>
                                <th scope="col">����ȯ��</th>
                                <th scope="col">������</th>
                                <th scope="col">��Ź�Ƶ�</th>
                                <th scope="col">�Ծ翩��</th>
                                <th scope="col">�Ծ�����</th>
                            </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                           family_support_yn = rs("family_support_yn")
						   family_disab = rs("family_disab")
						   family_merit = rs("family_merit")
						   family_serius = rs("family_serius")
						   family_pensioner = rs("family_pensioner")
						   family_witak = rs("family_witak")
						   family_holt = rs("family_holt")
						   if rs("family_holt_date") = "1900-01-01" then
						            family_holt_date = ""
							  else 
							        family_holt_date = rs("family_holt_date")
						   end if
	           			%>
							<tr>
                                <td><%=rs("family_rel")%>&nbsp;</td>
                                <td><%=rs("family_name")%>&nbsp;</td>
                                <td><%=rs("family_person1")%>-<%=rs("family_person2")%>&nbsp;</td>
                                <td><%=rs("family_national")%>&nbsp;</td>
                                <td><input type="checkbox" name="support_check" value="Y" <% if family_support_yn = "Y" then %>checked<% end if %> id="support_check"></td>
                                <td>
								<input type="checkbox" name="disab_check" value="Y" <% if family_disab = "Y" then %>checked<% end if %> id="disab_check"></td>
                                <td><input type="checkbox" name="merit_check" value="Y" <% if family_merit = "Y" then %>checked<% end if %> id="merit_check"></td>
                                <td><input type="checkbox" name="serius_check" value="Y" <% if family_serius = "Y" then %>checked<% end if %> id="serius_check"></td>
                                <td><input type="checkbox" name="pensioner_check" value="Y" <% if family_pensioner = "Y" then %>checked<% end if %> id="pensioner_check"></td>
                                <td><input type="checkbox" name="witak_check" value="Y" <% if family_witak = "Y" then %>checked<% end if %> id="witak_check"></td>
                                <td><input type="checkbox" name="holt_check" value="Y" <% if family_holt = "Y" then %>checked<% end if %> id="holt_check"></td>
                                <td><%=family_holt_date%>&nbsp;</td>
                                <td>&nbsp;</td>
                        <% if y_final <> "Y" then  %>
                                <td>
                                <a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=rs("family_empno")%>&family_seq=<%=rs("family_seq")%>&emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_family_add_pop','scrollbars=yes,width=750,height=370')">����</a></td>
                        <%    else  %>
                                <td>&nbsp;</td>
                        <% end if  %>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
              <% if y_final <> "Y" then  %>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=emp_no%>&emp_name=<%=emp_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=370')" class="btnType04">�ξ簡���߰�</a>
					</div>    
              <%   else  %>
                       <br><br>
			  <%   end if  %>                    
                    </td>
			      </tr>
				  </table>
           <h3 class="stit">�� �������� ���� ��Ϲ�� �� ���� ���� ��<br>&nbsp;<br>
                1. �������� �ʼ� ���<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �ҵ�����Ű� ������� �λ����>�������׿��� ������ �ҵ���� ������ Ȯ���� ��<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �⺻���� �Ǵ� Ư�������� ���� ���� �� �̵���ڴ� �ݵ�� ����� �ؾ� ��(��Ͻ� �ֹε�Ϲ�ȣ�� �ʼ� ������)<br>&nbsp;<br>
                2. �Ƿ��/��αݸ��� �ۼ����<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �ҵ�����Ű� �� �Ƿ��. ��α� ������ ����ϸ� �ش� ������ ����� �� �����Ƿ� ������ ���Ͽ� ������ �ʿ�� ����.<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �Ƿ�� �ۼ��� ����û �ڷ�� �Ƿ��� �ڷḦ �ߺ��Ͽ� �Է��ϸ� ���� ��¡����� �� �� ������ �ݵ�� Ȯ���Ͽ� ����<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� ��α��� ���ΰ� �⺻����������� ����� �� �ξ簡���� ������ ��α��� ���������.<br>&nbsp;<br>
                3. ����û�ݾ�/�׹��Ǳݾ� ���е��<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �����ڷ� ��Ͻ� ����û���� �߱޹��� �ڷ�� ����û�ݾ׿� �Է��ϰ�, �� �� �ڷ�� ��Ÿ�ݾ�(�� ���Ǳݾ�)�� �����Ͽ� ����ؾ� ��.<br>&nbsp;<br>
                4. �ſ�ī��,���ݿ�����.����ī�� ���<br>
                &nbsp;&nbsp;&nbsp;&nbsp;�� �ſ�ī��/�ñݿ�����/����ī�� �ݾ� �Է½� �Ϲ��հ�� ������� �� ��ر��� ��� �հ���� �����Ͽ� �Է�.</h3>
                <input type="hidden" name="family_empno" value="<%=in_empno%>" ID="Hidden1">  
                <input type="hidden" name="y_final" value="<%=y_final%>" ID="Hidden1">               
			</form>
		</div>				
	</div>        				
	</body>
</html>

