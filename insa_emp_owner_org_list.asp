<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_emp_owner_org_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	field_check=Request("field_check")
	field_bonbu=Request("field_bonbu")
	field_saupbu=Request("field_saupbu")
	field_team=Request("field_team")
	field_org=Request("field_org")
	view_c = Request("view_c")
  else
	view_condi=Request.form("view_condi")
	field_check=Request.form("field_check")
	field_bonbu=Request.form("field_bonbu")
	field_saupbu=Request.form("field_saupbu")
	field_team=Request.form("field_team")
	field_org=Request.form("field_org")
	view_c = Request.form("view_c")
End if

If view_condi = "" Then
	view_condi = "���̿��������"
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "ASC"
end if

order_Sql = " ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_reside_place,emp_no,emp_in_date " + view_sort

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "bonbu"
End If

If field_check = "total" Then
       owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"')"
	   field_check = ""
   else
       if view_c = "bonbu" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_bonbu like '%" + field_bonbu + "%')"
       end if
	   if view_c = "saupbu" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_saupbu like '%" + field_saupbu + "%')"
       end if
	   if view_c = "team" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_team like '%" + field_team + "%')"
       end if
	   if view_c = "orgm" Then
              owner_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000') and (emp_company = '"&view_condi&"') and (emp_org_name like '%" + field_org + "%')"
       end if
End If


Sql = "SELECT count(*) FROM emp_master " + owner_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_master " + owner_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " ���� ��Ȳ - ������������ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
			
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('bonbu1').style.display = '';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('orgm1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = '';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('orgm1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[2].checked")) {
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = '';
					document.getElementById('orgm1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[3].checked")) {
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('orgm1').style.display = '';
				}	
			}
			
			function insa_emp_owner_org(val, val2, val3) {

            if (!confirm("�ش������� ���� ���������� ���� �Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			
//			alert (val);
			
			document.frm.view_condi1.value = document.getElementById(val).value;
//			alert (val2);
//			document.frm.view_c1.value = document.getElementById(val2).value;
//			alert (val3);
//			document.frm.field_check1.value = document.getElementById(val3).value;
			
//		    alert (val3);
		
            document.frm.action = "insa_emp_owner_org_save.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_asses_promo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_owner_org_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>ȸ�� �˻�</dt>
                        <dd>
                            <p>
                               <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
								<input type="radio" name="view_c" value="bonbu" <% if view_c = "bonbu" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                ����
                                <input type="radio" name="view_c" value="saupbu" <% if view_c = "saupbu" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                �����
                                <input type="radio" name="view_c" value="team" <% if view_c = "team" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                ��
                                <input type="radio" name="view_c" value="orgm" <% if view_c = "orgm" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                �Ҽ�
								</label>
                                <label id="bonbu1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���� ��</strong>
                                	<input name="field_bonbu" type="text" value="<%=field_bonbu%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
								 <label id="saupbu1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����� ��</strong>
                                	<input name="field_saupbu" type="text" value="<%=field_saupbu%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
                                 <label id="team1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� ��</strong>
                                	<input name="field_team" type="text" value="<%=field_team%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
                                 <label id="orgm1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�Ҽ� ��</strong>
                                	<input name="field_org" type="text" value="<%=field_org%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                      
                
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="*" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">�����Ի���</th>
								<th scope="col">�Ҽӹ߷���</th>
								<th scope="col">����ó</th>
                                <th scope="col">����ȸ��</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <th scope="col">���</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
						if rs("emp_org_baldate") = "1900-01-01" then
						   emp_org_baldate = ""
						   else 
						   emp_org_baldate = rs("emp_org_baldate")
						end if
						if rs("emp_grade_date") = "1900-01-01" then
						   emp_grade_date = ""
						   else 
						   emp_grade_date = rs("emp_grade_date")
						end if
						%>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td><%=rs("emp_reside_company")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_emp_master_view.asp?view_condi=<%=rs("emp_company")%>&emp_no=<%=rs("emp_no")%>&u_type=<%=""%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=480')">��ȸ</a>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
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
				    <td width="20%">
					<div class="btnCenter">
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_emp_owner_org_list.asp?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&field_org=<%=field_org%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_emp_owner_org_list.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&field_org=<%=field_org%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_emp_owner_org_list.asp?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&field_org=<%=field_org%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_emp_owner_org_list.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&field_org=<%=field_org%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_emp_owner_org_list.asp?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&field_org=<%=field_org%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="insa_emp_owner_org('view_condi','view_c','field_check');return false;" class="btnType04">�������� ����</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
        <input type="hidden" name="user_id">
		<input type="hidden" name="pass">
        <input type="hidden" name="view_condi1" value="<%=view_condi%>" ID="Hidden1">
        <input type="hidden" name="view_c1" value="<%=view_c%>" ID="Hidden1">
        <input type="hidden" name="field_check1" value="<%=field_check%>" ID="Hidden1">
        <input type="hidden" name="field_bonbu1" value="<%=field_bonbu%>" ID="Hidden1">
        <input type="hidden" name="field_saupbu1" value="<%=field_saupbu%>" ID="Hidden1">
        <input type="hidden" name="field_team1" value="<%=field_team%>" ID="Hidden1">
        <input type="hidden" name="field_org1" value="<%=field_org%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

