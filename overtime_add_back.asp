<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")

mg_ce_id = user_id
mg_ce = user_name
apct_no = 0
company = ""
dept = ""
work_item = ""
from_hh = ""
from_mm = ""
to_hh = ""
to_mm = ""
work_gubun = ""
overtime_amt = 0
work_memo = ""
cancel_yn = "N"
end_sw = "N"
reg_id = user_id
reg_date = now()

curr_date = mid(cstr(now()),1,10)
work_date = mid(cstr(now()),1,10)

strNowWeek = WeekDay(work_date)
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

company = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "��Ư�� ���"
if u_type = "U" then

	work_date = request("work_date")
	mg_ce_id = request("mg_ce_id")

	sql = "select * from overtime where work_date = '" + work_date + "' and mg_ce_id = '" + mg_ce_id + "'"
	set rs = dbconn.execute(sql)

	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
	set rs_memb=dbconn.execute(sql)

	if	rs_memb.eof or rs_memb.bof then
		mg_ce = "ERROR"
	  else
		mg_ce = rs_memb("user_name")
	end if
	rs_memb.close()						

	if isnull(rs("acpt_no")) then
		acpt_no = 0
	  else
		acpt_no = rs("acpt_no")
	end if
	mg_ce_id = rs("mg_ce_id")
	company = rs("company")
	dept = rs("dept")
	work_item = rs("work_item")
	from_hh = mid(rs("from_time"),1,2)
	from_mm = mid(rs("from_time"),3,2)
	to_hh = mid(rs("to_time"),1,2)
	to_mm = mid(rs("to_time"),3,2)
	work_gubun = rs("work_gubun") + "/" + cstr(rs("overtime_amt"))
	overtime_amt = int(rs("overtime_amt"))
	work_memo = rs("work_memo")
	cancel_yn = rs("cancel_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	rs.close()

	title_line = "��Ư�� ����"
end if
if end_sw = "Y" then
	end_view = "����"
  else
  	end_view = "����"
end if

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
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
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
				if(document.frm.company.value =="" ) {
					alert('ȸ�縦 �����ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.work_item.value =="") {
					alert('�׸��� �����ϼ���');
					frm.work_item.focus();
					return false;}
				if(document.frm.work_gubun.value =="") {
					alert('�۾������� �����ϼ���');
					frm.work_gubun.focus();
					return false;}
				if(document.frm.work_date.value > document.frm.curr_date.value) {
					alert('�۾����� �����Ϻ��� Ů�ϴ�');
					frm.work_date.focus();
					return false;}			
			
				if(document.frm.from_hh.value >"23"||document.frm.from_hh.value <"00") {
					alert('���� �ð��� �߸��Ǿ����ϴ�');
					frm.from_hh.focus();
					return false;}
				if(document.frm.from_mm.value >"59"||document.frm.from_mm.value <"00") {
					alert('���� ���� �߸��Ǿ����ϴ�');
					frm.from_mm.focus();
					return false;}
				if(document.frm.to_hh.value >"23"||document.frm.to_hh.value <"00") {
					alert('���� �ð��� �߸��Ǿ����ϴ�');
					frm.to_hh.focus();
					return false;}
				if(document.frm.to_mm.value >"59"||document.frm.to_mm.value <"00") {
					alert('���� ���� �߸��Ǿ����ϴ�');
					frm.to_mm.focus();
					return false;}
			
				if(document.frm.to_hh.value < document.frm.from_hh.value) {
					alert('����ð��� ���۽ð� ���� �����ϴ�');
					frm.to_hh.focus();
					return false;}
			
				if(document.frm.from_hh.value == document.frm.to_hh.value) {
					if(document.frm.to_mm.value <= document.frm.from_mm.value) {
						alert('����ð��� ���۽ð� ���� �����ϴ�');
						frm.to_mm.focus();
						return false;}}
				
			
				if(document.frm.work_memo.value =="" ) {
					alert('�۾������� �Է��ϼ���');
					frm.work_memo.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function week_check() {
			
			a = document.frm.work_date.value.substring(0,4);
			b = document.frm.work_date.value.substring(5,7);
			c = document.frm.work_date.value.substring(8,10);
			
			var newDate = new Date(a,b-1,c); 
			var s = newDate.getDay(); 
			
			switch(s) {
				case 0: str = "�Ͽ���" ; break;
				case 1: str = "������" ; break;
				case 2: str = "ȭ����" ; break;
				case 3: str = "������" ; break;
				case 4: str = "�����" ; break;
				case 5: str = "�ݿ���" ; break;
				case 6: str = "�����" ; break;
				}
			
				document.frm.week.value = str;			
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_add_save.asp" method="post" name="frm">
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
								<th class="first">�۾���</th>
								<td class="left">
                                <input name="work_date" type="text" value="<%=work_date%>" style="width:70px" id="datepicker">&nbsp;<%=week%>
							<%  if u_type = "U" then	%>
                                <input name="old_date" type="hidden" value="<%=work_date%>">
                            <%	end if	%>
                                </td>
								<th>�۾���</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                                <input name="curr_date" type="hidden" id="curr_date" value="<%=now()%>">
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                                </td>
							</tr>
							<tr>
								<th class="first">ȸ���</th>
								<td class="left">
								  <%
                                        Sql="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' order by trade_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="company" id="select" style="width:150px">
                                    <option value="">����</option>
                                    <% 
                                        do until rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = company then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
                                    <%
                                        	rs_etc.movenext()  
                                        loop 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                </td>
								<th>�μ���</th>
								<td class="left"><input name="dept" type="text" id="dept" style="width:150px" notnull errname="������" onKeyUp="checklength(this,20)" value="<%=dept%>"></td>
							</tr>
							<tr>
								<th class="first">�۾��׸�</th>
								<td class="left">
							<%  if acpt_no > 0 then	%>
								<%=work_item%><input name="work_item" type="hidden" id="work_item" value="<%=work_item%>">
							<% 	  else	%>
                                <select name="work_item" id="work_item" style="width:150px">
                                    <option value="">����</option>
                                    <option value="����" <%If work_item = "����" then %>selected<% end if %>>����</option>
                                    <option value="������" <%If work_item = "������" then %>selected<% end if %>>������</option>
                                    <option value="��Ÿ" <%If work_item = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
							<%  end if	%>
            					</td>
								<th>�۾��ð�</th>
								<td class="left">
                                <input name="from_hh" type="text" id="from_hh" size="2" maxlength="2" value="<%=from_hh%>">��
								<input name="from_mm" type="text" id="from_mm" size="2" maxlength="2" value="<%=from_mm%>">�� ~
								<input name="to_hh" type="text" id="to_hh" size="2" maxlength="2" value="<%=to_hh%>">��
								<input name="to_mm" type="text" id="to_mm" size="2" maxlength="2" value="<%=to_mm%>">��
                                </td>
							</tr>
							<tr>
								<th class="first">�߱��׸�</th>
								<td class="left"><select name="work_gubun" id="work_gubun" style="width:150px">
								  <option value="">����</option>
								  <%
                                  Sql="select * from etc_code where etc_type = '41' order by etc_amt asc"
                                  rs_etc.Open Sql, Dbconn, 1
                                  do until rs_etc.eof
                                  		work_gubun_amt = rs_etc("etc_name") + "/" + cstr(rs_etc("etc_amt"))
                                  %>
								  <option value='<%=work_gubun_amt%>' <%If work_gubun_amt = work_gubun then %>selected<% end if %>><%=work_gubun_amt%></option>
								  <%
                                        rs_etc.movenext()
                                  loop
                                  rs_etc.close()						
                                  %>
						      </select></td>
								<th>�۾�����</th>
								<td class="left"><input name="work_memo" type="text" id="work_memo" notnull errname="�۾�����" onKeyUp="checklength(this,50)"  style="width:200px" value="<%=work_memo%>"></td>
							</tr>
							<tr id="cancel_col" style="display:none">
								<th class="first">��ҿ���</th>
								<td class="left">
								<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px" ID="Radio1">���
				                <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px" ID="Radio2">��û
								</td>
                                <th>��������</th>
								<td class="left"><%=end_view%><input name="end_sw" type="hidden" id="end_sw" value="<%=end_sw%>"></td>
							</tr>
							<tr id="info_col" style="display:none">
								<th class="first">�������</th>
								<td class="left">&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                                <th>��������</th>
								<td class="left">&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
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

