<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

be_pg = "insa_confirm_mg.asp"

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

from_date = request("from_date")
to_date = request("to_date")
cfm_type=Request("cfm_type")
company=Request("company")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
    cfm_type=Request.form("cfm_type")
    company=Request.form("company")
  else
	from_date = request("from_date")
    to_date = request("to_date")
    cfm_type=Request("cfm_type")
    company=Request("company")
end if

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	cfm_type = "��������"
	company = "��ü"
End If

if company = "��ü" then
	com_sql = ""
  else
  	com_sql = " (cfm_company ='"+company+"') and "
end if
if cfm_type = "��ü" then
	type_sql = ""
  else
  	type_sql = " (cfm_type ='"+cfm_type+"') and "
end if

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


Sql = "SELECT count(*) FROM emp_confirm where "+com_sql+type_sql+" cfm_date >= '"+from_date+"' and cfm_date <= '"+to_date+"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM emp_confirm where "+com_sql+type_sql+" cfm_date >= '"+from_date+"' and cfm_date <= '"+to_date+"' ORDER BY cfm_type,cfm_seq DESC limit "& stpage & "," &pgsize

Rs.Open Sql, Dbconn, 1

title_line = " ������ �߱� ��Ȳ "
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
				return "4 1";
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_welfare_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_confirm_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								<strong>�߱��� : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>  ��  </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>ȸ��</strong>
                                <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="company" id="company" type="text" style="width:150px">

                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								loop
								rs_org.Close()
							  %>
            					</select>
                                </label>
								<strong>����������</strong>
                                <select name="cfm_type" id="cfm_type" style="width:100px">
                                    <option value="��ü" <%If cfm_type = "��ü" then %>selected<% end if %>>��ü</option>
                                    <option value="��������" <%If cfm_type = "��������" then %>selected<% end if %>>��������</option>
                                    <option value="�������" <%If cfm_type = "�������" then %>selected<% end if %>>�������</option>
                                    <option value="��õ¡��" <%If cfm_type = "��õ¡��" then %>selected<% end if %>>��õ¡��</option>
                                    <option value="���ټ���" <%If cfm_type = "���ټ���" then %>selected<% end if %>>���ټ���</option>
                                </select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="*" >
						</colgroup>
						<thead>
						  <tr>
							<th class="first" scope="col">���</th>
							<th scope="col">����</th>
							<th scope="col">����</th>
							<th scope="col">��å</th>
							<th scope="col">ȸ��</th>
                            <th scope="col">�Ҽ�</th>
                            <th scope="col">�߱���</th>
							<th scope="col">�߱޹�ȣ</th>
							<th scope="col">������</th>
							<th scope="col">�뵵</th>
                            <th scope="col">���ó</th>
							<th scope="col">�ֹι�ȣ</th>
                            <th scope="col">���</th>
						  </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

		                  cfm_empno = rs("cfm_empno")

                         if cfm_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&cfm_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_job = Rs_emp("emp_job")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if
	           			%>
							<tr>
								<td class="first"><%=rs("cfm_empno")%></td>
                                <td>
								 <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("cfm_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("cfm_emp_name")%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("cfm_company")%>&nbsp;</td>
                                <td><%=rs("cfm_org_name")%>&nbsp;</td>
                                <td><%=rs("cfm_date")%>&nbsp;</td>
                                <td>��&nbsp;<%=rs("cfm_number")%>-<%=rs("cfm_seq")%>&nbsp;ȣ</td>
                                <td><%=rs("cfm_type")%>&nbsp;</td>
                                <td><%=rs("cfm_use")%>&nbsp;</td>
								<td><%=rs("cfm_use_dept")%>&nbsp;</td>
                                <td><%=rs("cfm_person1")%>-<%=rs("cfm_person2")%>&nbsp;</td>
                                <td><%=rs("cfm_comment")%>&nbsp;</td>
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
                    <a href="/insa_excel_cfmlist.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_confirm_mg.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_confirm_mg.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_confirm_mg.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_confirm_mg.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_confirm_mg.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

