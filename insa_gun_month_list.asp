<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

curr_date = mid(cstr(now()),1,10)

emp_no = request("emp_no")
be_pg = request("be_pg")
page = request("page")
page_cnt = request("page_cnt")

be_pg1 = "insa_gun_month_list.asp"
'be_pg = "insa_gun_month_list.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

ck_sw=Request("ck_sw")
win_sw = "close"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_empno")  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
   if not Rs_emp.eof then
      in_name = rs_emp("emp_name")
	  else
      response.write"<script language=javascript>"
	  response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
	  response.write"</script>"
	  Response.End	
   end if
   rs_emp.close()
End If

sql = "select * from emp_master where emp_no = '" + in_empno + "'"
Rs.Open Sql, Dbconn, 1


title_line = " ���� ������Ȳ(������)....."

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
				return "3 1";
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if (document.frm.in_empno.value == "") {
					alert ("����� �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}
				
				return true;
			}
            function s_sinchung(val, val2, val3, val4, val5) {

            if (!confirm("�������� ��û�Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;
			document.frm.in_name.value = val3;
			document.frm.in_name.value = val5;

            if (document.getElementById(val3).value == "")
            { alert("�������� �Է����ּ���!"); return; }
			if (document.getElementById(val4).value == "")
            { alert("���������� �������ּ���!"); return; }

            document.frm.cfm_use.value = document.getElementById(val4).value;
            document.frm.action = "insa_resign_print.asp";
            document.frm.submit();
            }	
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_gun_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_gun_month_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�����ް� �ϼ�</dt>
                        <dd>
							<strong>�⵵: </strong>
								<label>
        						<input name="in_year" type="text" id="in_year" value="<%=year_year%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>���������: </strong>
                                <label>
                               	<input name="in_yuncha_date" type="text" id="in_yuncha_date" value="<%=year_yuncha_date%>" readonly="true" style="width:60px; text-align:left">
								</label>
                            <strong>�ټӳ��: </strong>
                                <label>
                               	<input name="in_continu_year" type="text" id="in_continu_year" value="<%=year_continu_year%>" readonly="true" style="width:40px; text-align:left">
                                -
                                <input name="in_continu_month" type="text" id="in_continu_month" value="<%=year_continu_month%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>�߻�����: </strong>
                                <label>
                               	<input name="in_basic_count" type="text" id="in_basic_count" value="<%=year_basic_count%>" readonly="true" style="width:40px; text-align:left">
                                -
                                <input name="in_add_count" type="text" id="in_add_count" value="<%=year_add_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>��뿬��: </strong>
                                <label>
                               	<input name="in_use_count" type="text" id="in_use_count" value="<%=year_use_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>�ܿ�����: </strong>
                                <label>
                               	<input name="in_remain_count" type="text" id="in_remain_count" value="<%=year_remain_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="*%" >
						</colgroup>
						<thead>
						    <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-left:1px solid #e3e3e3;">���</th>
                                <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;��</th>
				                <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
			                </tr>
                            <tr>
								<th class="first" scope="col" style=" border-right:1px solid #e3e3e3;">����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">�����ް�</th>
                                <th scope="col">�ð��ܱٹ�</th>
                                <th scope="col">���ϱٹ�</th>
                                <th scope="col">�ܱ�</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                                <th scope="col">���</th>
                                <th scope="col">��Ÿ</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
                                <td class="first">
                                <a href="#" onClick="pop_Window('insa_gun_month_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','gun_monthview','scrollbars=yes,width=800,height=400')"><%=rs("emp_no")%></a>
								</td>
                                <td>
                                <a href="insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("emp_company")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td>
								 <input name="end_date" type="text" size="10" readonly="true" id="datepicker" style="width:60px;">&nbsp;</td>
                                <td>
                                <select name="end_type" id="end_type" value="<%=end_type%>" style="width:90px">
			            	        <option value="" <% if end_type = "" then %>selected<% end if %>>����</option>
				                    <option value='ȸ�����' <%If end_type = "ȸ�����" then %>selected<% end if %>>ȸ�����</option>
                                    <option value='������' <%If end_type = "������" then %>selected<% end if %>>������</option>
                                    <option value='���λ���' <%If end_type = "���λ���" then %>selected<% end if %>>���λ���</option>
                                    <option value='¡��' <%If end_type = "¡��" then %>selected<% end if %>>¡��</option>
                                    <option value='����' <%If end_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='����' <%If end_type = "����" then %>selected<% end if %>>����</option>
                                    <option value='ġ��' <%If end_type = "ġ��" then %>selected<% end if %>>ġ��</option>
                                </select>                                 
                                </td>
                                <td class="left">
								<input name="end_comment" type="text" id="end_comment" style="width:120px" onKeyUp="checklength(this,30)" value="<%=cfm_comment%>">
                                </td>                                
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
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
                       <div class="btnRight">
                         <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                       </div>
                    </td>
			     </tr>
			    </table>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

