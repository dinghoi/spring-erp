<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
acpt_no = request("acpt_no")
be_pg = request("be_pg")

page = request("page")
from_date = request("from_date")
to_date = request("to_date")
date_sw = request("date_sw")
process_sw = request("process_sw")
field_check = request("field_check")
field_view = request("field_view")
view_sort = request("view_sort")
page_cnt = request("page_cnt")
condi_com = request("company")
view_c = request("view_c")

visit_date = mid(now(),1,10)

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)

acpt_date = mid(cstr(rs("acpt_date")),1,10)
acpt_hh = int(datepart("h",rs("acpt_date")))
acpt_mm = int(datepart("n",rs("acpt_date")))
acpt_ss = datepart("s",rs("acpt_date"))

if isnull(rs("dev_inst_cnt")) or rs("dev_inst_cnt") = "" then
	dev_inst_cnt = "1"
  else
  	dev_inst_cnt = rs("dev_inst_cnt")
end if

title_line = "�뷮�� �Ϸ� ���"
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
												$( "#datepicker" ).datepicker("setDate", "<%=visit_date%>" );
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

				if (document.frm.juso_mod_ck.checked == true) {
					if(document.frm.sido.value == "" || document.frm.gugun.value == "" || document.frm.dong.value == "" || document.frm.addr.value == "") {
						alert('�ּҸ� �Է��ϼ���');
						frm.juso_mod.focus();
						return false;}
				}
				if(document.frm.dev_inst_cnt.value < 0 || document.frm.dev_inst_cnt.value > 999) {
					alert('��ġ����� 999���� ũ�ų� �߸��Ǿ����ϴ�');
					frm.dev_inst_cnt1.focus();
					return false;}
				if(document.frm.ran_cnt.value < 0 || document.frm.ran_cnt.value > 999) {
					alert('�������� 999���� ũ�ų� �߸��Ǿ����ϴ�');
					frm.ran_cnt.focus();
					return false;}
				if(document.frm.work_man_cnt.value < 1 || document.frm.work_man_cnt.value > 30) {
					alert('�۾� �ο��� 30���� ũ�ų� �߸��Ǿ����ϴ�');
					frm.work_man_cnt.focus();
					return false;}
				if(document.frm.alba_cnt.value < 0 || document.frm.alba_cnt.value > 30) {
					alert('�˹� �ο��� 30���� ũ�ų� �߸��Ǿ����ϴ�');
					frm.alba_cnt.focus();
					return false;}
				if(document.frm.visit_date.value == "") {
					alert('ó�����ڸ� �Է��ϼ���!!');
					frm.visit_date.focus();
					return false;}
				if(document.frm.visit_hh.value >"23"||document.frm.visit_hh.value <"00") {
					alert('ó���ð��� �߸��Ǿ����ϴ�');
					frm.visit_hh.focus();
					return false;}
				if(document.frm.visit_mm.value >"59"||document.frm.visit_mm.value <"00") {
					alert('ó������ �߸��Ǿ����ϴ�');
					frm.visit_mm.focus();
					return false;}
				if(document.frm.visit_date.value < document.frm.acpt_date.value) {
					alert('�Ϸ����� �����Ϻ��� �����ϴ�');
					frm.visit_date.focus();
					return false;}
				if(document.frm.visit_date.value > document.frm.curr_date.value) {
					alert('�Ϸ����� �����Ϻ��� �����ϴ�');
					frm.visit_date.focus();
					return false;}
				if(document.frm.visit_date.value == document.frm.acpt_date.value) {
					if(document.frm.visit_hh.value <= document.frm.acpt_hh.value) {
						alert('�Ϸ�ð��� �����ð� ���� �����ϴ�');
						frm.visit_hh.focus();
						return false;}}
				if(document.frm.visit_date.value == document.frm.acpt_date.value) {
					if(document.frm.visit_hh.value == document.frm.acpt_hh.value) {
						if(document.frm.visit_mm.value <= document.frm.acpt_mm.value) {
							alert('�Ϸ���� ������ ���� �����ϴ�');
							frm.visit_mm.focus();
							return false;}}}
				if(document.frm.att_file1.value =="" && document.frm.att_file2.value =="" && document.frm.att_file3.value =="" && document.frm.att_file4.value =="" && document.frm.att_file5.value =="") {
					alert('���� ÷�ΰ� ���� �ʾҽ��ϴ�');
					frm.att_file1.focus();
					return false;}

				{
				a=confirm('�Ϸ� ����� �Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function juso_mod_view() 
			{
				if (document.frm.juso_mod_ck.checked == true) {
					document.getElementById('juso_mod').style.display = ''; 
					document.getElementById('juso_mod_field').style.display = ''; }
				if (document.frm.juso_mod_ck.checked == false) {
					document.getElementById('juso_mod').style.display = 'none'; 
					document.getElementById('juso_mod_field').style.display = 'none'; }
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="large_result_reg_ok.asp" method="post" name="frm" enctype="multipart/form-data">
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
								<th class="first">����</th>
								<td class="left"><%=rs("company")%>&nbsp;<%=rs("dept")%>&nbsp;<%=rs("acpt_user")%>
                                <input name="acpt_no" type="hidden" id="acpt_no" value="<%=acpt_no%>">
								</td>
								<th>ó������</th>
								<td class="left"><%=rs("as_type")%></td>
                                </td>
							</tr>
							<tr>
								<th class="first">�����ּ�</th>
								<td colspan="3" class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%>&nbsp;&nbsp;&nbsp;&nbsp;<strong>������ ���ϸ� �����ϼ���</strong>
                                  <input name="juso_mod_ck" type="checkbox" id="juso_mod_ck" value="1"  onClick="juso_mod_view()">
                                <a href="#" class="btnType03" onclick="javascript:pop_area()" id="juso_mod" style="display:none">�ּҺ���</a></td>              
							</tr>
							<tr id="juso_mod_field" style="display:none">
								<th class="first">�����ּ�</th>
								<td colspan="3" class="left">
                                <input name="sido" type="text" id="sido" style="width:50px" readonly="true">
              					<input name="gugun" type="text" id="gugun" style="width:100px" readonly="true">
              					<input name="dong" type="text" id="dong" style="width:100px" readonly="true">
              					<input name="addr" type="text" id="addr" style="width:200px" onKeyUp="checklength(this,50)">
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id">
                                <input name="mg_ce" type="hidden" id="mg_ce">
                                <input name="reside_place" type="hidden" id="reside_place">
                                <input name="team" type="hidden" id="team">
                                </td>              
							</tr>
							<tr>
								<th class="first">ó���Ǽ�</th>
								<td colspan="3" class="left">
                                ��ġ���
                                <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);"  maxlength="3" value="<%=dev_inst_cnt%>">��&nbsp; 
                                ������
                                <input name="ran_cnt" type="text" id="ran_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="3">��&nbsp;
                                �۾��η�
                                <input name="work_man_cnt" type="text" id="work_man_cnt" style="width:30px;text-align:right" value="1" maxlength="2" readonly="true">��&nbsp;
                                �˹��ο�
                                <input name="alba_cnt" type="text" id="alba_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="2">��
								<a href="#" id="work_ce" class="btnType03" onClick="pop_Window('work_ce_add.asp?acpt_no=<%=rs("acpt_no")%>','work_ce_add_pop','scrollbars=yes,width=700,height=500')">2���̻��۾��ηµ��</a>
                                </td>
							</tr>
							<tr>
								<th class="first">ó������</th>
								<td colspan="3" class="left">
                                <input name="visit_date" type="text" style="width:70px" id="datepicker">
                                <input name="visit_hh" type="text" id="visit_hh" size="2" maxlength="2"><strong>��</strong>                                
                                <input name="visit_mm" type="text" id="visit_mm" size="2" maxlength="2"><strong>��</strong>
                                </td>              
							</tr>
							<tr>
							  <th class="first">ó������</th>
							  <td colspan="3" class="left"><textarea name="as_history" rows="2" id="textarea"><%=rs("as_history")%></textarea></td>
					      </tr>
							<tr>
								<th class="first">÷������1</th>
								<td class="left" colspan="3"><input name="att_file1" type="file" id="att_file1" size="70"></td>
							</tr>
							<tr>
								<th class="first">÷������2</th>
								<td class="left" colspan="3"><input name="att_file2" type="file" id="att_file2" size="70"></td>
							</tr>
							<tr>
								<th class="first">÷������3</th>
								<td class="left" colspan="3"><input name="att_file3" type="file" id="att_file3" size="70"></td>
							</tr>
							<tr>
								<th class="first">÷������4</th>
								<td class="left" colspan="3"><input name="att_file4" type="file" id="att_file4" size="70"></td>
							</tr>
							<tr>
								<th class="first">÷������5</th>
								<td class="left" colspan="3"><input name="att_file5" type="file" id="att_file5" size="70"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                <input name="page" type="hidden" id="page" value="<%=page%>">
                <input name="from_date" type="hidden" id="from_date" value="<%=from_date%>">
                <input name="acpt_date" type="hidden" id="acpt_date" value="<%=acpt_date%>">
                <input name="acpt_hh" type="hidden" id="acpt_hh" value="<%=acpt_hh%>">
                <input name="acpt_mm" type="hidden" id="acpt_mm" value="<%=acpt_mm%>">
                <input name="curr_date" type="hidden" id="curr_date" value="<%=curr_date%>">
                <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                <input name="date_sw" type="hidden" id="date_sw" value="<%=date_sw%>">
                <input name="process_sw" type="hidden" id="process_sw" value="<%=process_sw%>">
                <input name="field_check" type="hidden" id="field_check" value="<%=field_check%>">
                <input name="field_view" type="hidden" id="field_view" value="<%=field_view%>">
                <input name="view_sort" type="hidden" id="view_sort" value="<%=view_sort%>">
                <input name="condi_com" type="hidden" id="condi_com" value="<%=condi_com%>">
                <input name="view_c" type="hidden" id="view_c" value="<%=view_c%>">
                <input name="company" type="hidden" id="company" value="<%=rs("company")%>">
                <input name="dept" type="hidden" id="dept" value="<%=rs("dept")%>">
                <input name="as_type" type="hidden" id="as_type" value="<%=rs("as_type")%>">
                <input name="o_sido" type="hidden" id="o_sido" value="<%=rs("sido")%>">
                <input name="o_gugun" type="hidden" id="o_gugun" value="<%=rs("gugun")%>">
                <input name="o_dong" type="hidden" id="o_dong" value="<%=rs("dong")%>">
                <input name="o_addr" type="hidden" id="o_addr" value="<%=rs("addr")%>">
			</form>
		</div>				
	</body>
</html>

