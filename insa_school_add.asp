<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
sch_empno = request("sch_empno")
sch_seq = request("sch_seq")
emp_name = request("emp_name")

sch_start_date = ""
sch_end_date = ""
sch_school_name = ""
sch_dept = ""
sch_major = ""
sch_sub_major = ""
sch_degree = ""
sch_finish = ""
sch_comment = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " �з»��� ��� "
if u_type = "U" then

	Sql="select * from emp_school where sch_empno = '"&sch_empno&"' and sch_seq = '"&sch_seq&"'"
	Set rs=DbConn.Execute(Sql)

    sch_empno = rs("sch_empno")
    sch_seq = rs("sch_seq")
	sch_start_date = rs("sch_start_date")
    sch_end_date = rs("sch_end_date")
    sch_dept = rs("sch_dept")
    sch_major = rs("sch_major")
    sch_sub_major = rs("sch_sub_major")
    sch_degree = rs("sch_degree")
	sch_finish = rs("sch_finish")
    view_condi = rs("sch_comment")
	if view_condi = "1" then 
	        sch_school_name = rs("sch_school_name")
	   else
	        sch_school_name = rs("sch_school_name")
	end if
	
	rs.close()

	title_line = " �з»��� ���� "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=sch_start_date%>" );
			});	 
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=sch_end_date%>" );
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
				if(document.frm.sch_end_date.value =="") {
					alert('�������ڸ� �Է��ϼ���');
					frm.sch_end_date.focus();
					return false;}
				if(document.frm.view_condi.value =="1") 
					if(document.frm.sch_high_name.value =="") {
						alert('�б����� �Է��ϼ���');
						frm.sch_high_name.focus();
						return false;}	
				if(document.frm.view_condi.value =="2") 
					if(document.frm.sch_school_name.value =="") {
						alert('�б����� �����ϼ���');
						frm.sch_school_name.focus();
						return false;}	
			    if(document.frm.sch_finish.value =="") {
					alert('�������θ� �����ϼ���');
					frm.sch_finish.focus();
					return false;}
				if(document.frm.sch_dept.value =="") {
					alert('�а��� �Է��ϼ���');
					frm.sch_dept.focus();
					return false;}
				if(document.frm.sch_major.value =="") {
					alert('������ �Է��ϼ���');
					frm.sch_major.focus();
					return false;}
				
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')   
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function condi_view() {
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.view_condi[" + j + "].checked")) {
						k = j + 1
					}
				}
				if (k==1){
					document.frm.sch_high_name.style.display = '';				
					document.frm.sch_school_name.style.display = 'none';				
				}
				if (k==2){
					document.frm.sch_high_name.style.display = 'none';				
					document.frm.sch_school_name.style.display = '';				
				}
			}			
				
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_school_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">���</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="sch_empno" type="text" id="sch_empno" size="14" value="<%=sch_empno%>" readonly="true">
                      <input type="hidden" name="sch_seq" value="<%=sch_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>�Ⱓ</th>
                      <td colspan="5" class="left">
                      <input name="sch_start_date" type="text" value="<%=sch_start_date%>" style="width:80px;text-align:center" id="datepicker">
                      &nbsp;-&nbsp;
                      <input name="sch_end_date" type="text" value="<%=sch_end_date%>" style="width:80px;text-align:center" id="datepicker1">
                      </td>
                    </tr>
                    <tr>  
                      <th>�б���</th>
                      <td colspan="5" class="left">
                      <input type="radio" name="view_condi" value="1" <% if view_condi = "1" then %>checked<% end if %> title="����б�" style="width:30px" onClick="condi_view()">����б�
                      <% if view_condi = "1" then %>
                           <input name="sch_high_name" type="text" id="sch_high_name" value="<%=sch_school_name%>" style="width:150px">
                      <%       else  %>    
                           <input name="sch_high_name" type="text" id="sch_high_name" style="display:none; width:150px"> 
                      <% end if %>   
                       <input type="radio" name="view_condi" value="2" <% if view_condi = "2" then %>checked<% end if %> title="����" style="width:30px" onClick="condi_view()">����
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '20' order by emp_etc_name asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
                    <% if view_condi = "2" then %>
					  <select name="sch_school_name" id="sch_school_name" value="<%=sch_school_name%>" style="width:150px">
                    <%       else  %>   
                      <select name="sch_school_name" id="sch_school_name" style="display:none; width:150px">
                    <% end if %>   
                         <option value="" <% if sch_school_name = "" then %>selected<% end if %>>����</option>
                			  <% 
								do until rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If sch_school_name = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							  %>
            		  </select> 
                      </td>     
                    </tr>
                    <tr>
                      <th>�а�</th>
                      <td class="left">
					  <input name="sch_dept" type="text" id="sch_dept" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_dept%>">&nbsp;</td>
					  <th>����</th>
                      <td class="left">
					  <input name="sch_major" type="text" id="sch_major" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_major%>">&nbsp;</td>
                      <th>������</th>
                      <td class="left">
					  <input name="sch_sub_major" type="text" id="sch_sub_major" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_sub_major%>">&nbsp;</td>
                    </tr>
                    <tr>  
                      <th>����</th>
                      <td class="left">
					  <select name="sch_finish" id="sch_finish" value="<%=sch_finish%>" style="width:100px">
				          <option value="" <% if sch_finish = "" then %>selected<% end if %>>����</option>
				          <option value='����' <%If sch_finish = "����" then %>selected<% end if %>>����</option>
				          <option value='����' <%If sch_finish = "����" then %>selected<% end if %>>����</option>
				          <option value='����' <%If sch_finish = "����" then %>selected<% end if %>>����</option>
                     </select>
                      <th>����</th>
                      <td colspan="3" class="left">                      
                      <select name="sch_degree" id="sch_degree" value="<%=sch_degree%>" style="width:100px">
				          <option value="" <% if sch_degree = "" then %>selected<% end if %>>����</option>
				          <option value='�����л�' <%If sch_degree = "�����л�" then %>selected<% end if %>>�����л�</option>
				          <option value='�л�' <%If sch_degree = "�л�" then %>selected<% end if %>>�л�</option>
				          <option value='����' <%If sch_degree = "����" then %>selected<% end if %>>����</option>
                          <option value='�ڻ�' <%If sch_degree = "�ڻ�" then %>selected<% end if %>>�ڻ�</option>
                     </select>
					  </td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

