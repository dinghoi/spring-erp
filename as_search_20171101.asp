<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
Dim strNowWeek

work_date = Request.form("work_date")
work_item = Request.form("work_item")
view_c = Request.form("view_c")
acpt_no = Request.form("acpt_no")
if work_date = "" then
	work_date = mid(cstr(now()),1,10)
	work_item = ""
	acpt_no = ""
	view_c = "acpt"
end if

if view_c = "acpt" then
	SQL = "select * from as_acpt where overtime <> 'Y' and ( acpt_no = '"&acpt_no&"') and as_process = '�Ϸ�' ORDER BY acpt_no ASC"
  else
	SQL = "select * from as_acpt where overtime <> 'Y' and ( as_type = '"&work_item&"') and visit_date = '"&work_date&"' and as_process = '�Ϸ�' ORDER BY acpt_no ASC"
end if
Rs.open SQL, Dbconn, 1

title_line = "A/S �˻�"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>CE �˻�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
			});	  
			function as_list(acpt_no,company,dept,work_date,work_item,dev_inst_cnt,ran_cnt,work_man_cnt,week,holi_sw)
			{
				opener.document.frm.acpt_no.value = acpt_no;
				opener.document.frm.company.value = company;
				opener.document.frm.dept.value = dept;
				opener.document.frm.work_date.value = work_date;
				opener.document.frm.work_item.value = work_item;
				opener.document.frm.dev_inst_cnt.value = dev_inst_cnt;
				opener.document.frm.ran_cnt.value = ran_cnt;
				opener.document.frm.work_man_cnt.value = work_man_cnt;
				opener.document.frm.week.value = week;
				opener.document.frm.holi_sw.value = holi_sw;
				window.close();
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
//				if(document.frm.work_item.value =="") {
//					alert('�۾������� �����ϼ���');
//					frm.work_item.focus();
//					return false;}
				if(document.frm.end_date.value >= document.frm.work_date.value) {
					alert('������ �����Դϴ�');
					frm.work_date.focus();
					return false;}
				{
					return true;
				}
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('work1').style.display = 'none';
					document.getElementById('work2').style.display = 'none';
					document.getElementById('acpt1').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('work1').style.display = '';
					document.getElementById('work2').style.display = '';
					document.getElementById('acpt1').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
 								<label>
                              	<input type="radio" name="view_c" value="acpt" <% if view_c = "acpt" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                ������ȣ
                                <input type="radio" name="view_c" value="work" <% if view_c = "work" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                �۾�����
								</label>
                                <label id="work1">
								<strong>�۾�����</strong>
                                <select name="work_item" id="work_item" style="width:100px">
                              		<option value="">����</option>
								    <option value="�湮ó��" <%If work_item = "�湮ó��" then %>selected<% end if %>>�湮ó��</option>
								    <option value="����ó��" <%If work_item = "����ó��" then %>selected<% end if %>>����ó��</option>
								    <option value="�űԼ�ġ" <%If work_item = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
								    <option value="�űԼ�ġ����" <%If work_item = "�űԼ�ġ����" then %>selected<% end if %>>�űԼ�ġ����</option>
								    <option value="������ġ" <%If work_item = "������ġ" then %>selected<% end if %>>������ġ</option>
								    <option value="������ġ����" <%If work_item = "������ġ����" then %>selected<% end if %>>������ġ����</option>
								    <option value="������" <%If work_item = "������" then %>selected<% end if %>>������</option>
								    <option value="����������" <%If work_item = "����������" then %>selected<% end if %>>����������</option>
								    <option value="���ȸ��" <%If work_item = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
								    <option value="��������" <%If work_item = "��������" then %>selected<% end if %>>��������</option>
								    <option value="��Ư��" <%If work_item = "��Ư��" then %>selected<% end if %>>��Ư��</option>
								    <option value="��Ÿ" <%If work_item = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
								</label>
								<label id="work2">
								<strong>�۾�����</strong>
                                	<input name="work_date" type="text" value="<%=work_date%>" style="width:70px" id="datepicker" id="work_date">
								</label>
								<label id="acpt1">
								<strong>������ȣ</strong>
                                	<input name="acpt_no" type="text" value="<%=acpt_no%>" style="width:70px" id="acpt_no">
								</label>
								<strong>���� : </strong><%=end_date%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="22%" >
							<col width="*" >
							<col width="20%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">������ȣ</th>
								<th scope="col">ȸ��</th>
								<th scope="col">������</th>
								<th scope="col">�����</th>
								<th scope="col">��ġ</th>
								<th scope="col">����</th>
								<th scope="col">�ο�</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							strNowWeek = WeekDay(rs("visit_date"))
							Select Case (strNowWeek)
							   Case 1
								   week = "�Ͽ���"
							   Case 2
								   week = "������"
							   Case 3
								   week = "ȭ����"
							   Case 4
								   week = "������"
							   Case 5
								   week = "�����"
							   Case 6
								   week = "�ݿ���"
							   Case 7
								   week = "�����"
							End Select
							if week = "�����" or week = "�Ͽ���" then
								holi_sw = "Y"
							  else	
								holi_sw = "N"
								sql = "select * from holiday where holiday = '" + cstr(rs("visit_date"))  + "'"
								set rs_etc = dbconn.execute(sql)
								if rs_etc.eof or rs_etc.bof then
									holi_sw = "N"
								  else
									holi_sw = "Y"
								end if
								rs_etc.close()
							end if
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="as_list('<%=rs("acpt_no")%>','<%=rs("company")%>','<%=rs("dept")%>','<%=rs("visit_date")%>','<%=rs("as_type")%>','<%=rs("dev_inst_cnt")%>','<%=rs("ran_cnt")%>','<%=rs("work_man_cnt")%>','<%=week%>','<%=holi_sw%>');"><%=rs("acpt_no")%></a>
                                </td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("mg_ce")%>&nbsp;(<%=rs("mg_ce_id")%>)</td>
								<td><%=rs("dev_inst_cnt")%></td>
								<td><%=rs("ran_cnt")%></td>
								<td><%=rs("work_man_cnt")%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="7">������ �����ϴ�</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>				
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="end_date">
			</form>
		</div>        				
	</body>
</html>

