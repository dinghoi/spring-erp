<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

view_condi = request("view_condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
  else
	view_condi = request("view_condi")
	owner_view=request("owner_view")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	ck_sw = "n"
end if


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi <> "" then
     if owner_view = "C" then  
	     Sql= "select * " & _
	          "    from emp_language a, emp_master b " & _
	          "    where a.lang_empno = b.emp_no AND b.emp_name like '%" + view_condi + "%' " & _
		      "    ORDER BY lang_empno,lang_seq ASC" 
       else
	     sql = "select * from emp_language where lang_empno = '"+view_condi+"' ORDER BY lang_empno,lang_seq ASC"
     end if
	 Rs.Open Sql, Dbconn, 1
end if
'Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " ���дɷ� ���� "
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
				return "1 1";
			}
			function goAction () {
			   window.close () ;
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("������ �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
			function language_del(val, val2, val3, val4) {

            if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			document.frm.lang_empno.value = val;
			document.frm.lang_seq.value = val2;
			document.frm.lang_empname.value = val3;
			document.frm.owner_view.value = val4;
		
            document.frm.action = "insa_language_del.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_language_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>������ �˻���</dt>
                        <dd>
                            <p>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">���
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">����
                                </label>
							<strong>���� : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
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
                            <col width="6%" >
                            <col width="11%" >
                            <col width="*" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>���</th>
                                <th>����</th>
                                <th>�Ҽ�</th>
                                <th>���б���</th>
                                <th>��������</th>
                                <th>����</th>
                                <th>�޼�</th>
                                <th>�����</th>
                                <th>����</th>
                                <th>����</th>
                                <th>���</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
						      lang_empno = rs("lang_empno")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&lang_empno&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
                                   emp_bonbu = rs_emp("emp_bonbu")
                                   emp_saupbu = rs_emp("emp_saupbu")
                                   emp_team = rs_emp("emp_team")
                                   emp_org_code = rs_emp("emp_org_code")
                                   emp_org_name = rs_emp("emp_org_name")
							  end if
							  rs_emp.close()							
						%>
							<tr>
                              <td><%=rs("lang_empno")%>&nbsp;</td>
                              <td><%=emp_name%>&nbsp;</td>
                              <td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td><%=rs("lang_id")%>&nbsp;</td>
                              <td><%=rs("lang_id_type")%>&nbsp;</td>
                              <td><%=rs("lang_point")%>&nbsp;</td>
                              <td><%=rs("lang_grade")%>&nbsp;</td>
                              <td><%=rs("lang_get_date")%>&nbsp;</td>
                              <td><a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=rs("lang_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')">���</a></td>
							  <td><a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=rs("lang_empno")%>&lang_seq=<%=rs("lang_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')">����</a></td>
                         <% if insa_grade = "0" then %>     
                              <td>
                              <a href="#" onClick="language_del('<%=rs("lang_empno")%>', '<%=rs("lang_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">����</a></td>
                         <%     else %>
                              <td>&nbsp;</td>
                         <% end if %>                                          
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						end if
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<% if owner_view = "T" then 
                              emp_no = view_condi
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()
				    %>
                    <a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=view_condi%>&emp_name=<%=emp_name%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">���л��� ���</a>
					<% end if %>
                    </div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="lang_empno" value="<%=lang_empno%>" ID="Hidden1">
                  <input type="hidden" name="lang_seq" value="<%=lang_seq%>" ID="Hidden1">
                  <input type="hidden" name="lang_empname" value="<%=lang_empname%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

