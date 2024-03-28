<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next

'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim be_pg, cfm_use, cfm_use_dept, cfm_comment, view_condi, ck_sw
Dim owner_view, title_line, rsConf
Dim in_empno, in_name, emp_name

be_pg = "/insa/insa_confirm_career_list.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

view_condi = Request("view_condi")
ck_sw = Request("ck_sw")

title_line = " ������ ������� �߱� "

if ck_sw = "n" then
	owner_view = Request.form("owner_view")
	view_condi = Request.form("view_condi")
  else
	owner_view = Request("owner_view")
	view_condi = Request("view_condi")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	ck_sw = "n"
end if

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_end_date, emp_company, emp_org_name, "
objBuilder.Append "	eomt.org_company, eomt.org_name "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_no < '900000' AND emp_end_date > '1900-01-01' "

if view_condi <> "" then
	if owner_view = "C" then
	'sql = "select * from emp_master where emp_name like '%"+view_condi+"%' and (emp_no < '900000') and (emp_end_date > '1900-01-01') ORDER BY emp_no,emp_company,emp_bonbu,emp_saupbu,emp_team ASC"
		objBuilder.Append "AND emp_name LIKE '%"&view_condi&"%' "
	else
		'sql = "select * from emp_master where emp_no = '"+view_condi+"' and (emp_no < '900000') and (emp_end_date > '1900-01-01') ORDER BY emp_no,emp_company,emp_bonbu,emp_saupbu,emp_team ASC"
		objBuilder.Append "AND emp_no = '"&view_condi&"' "
	end If

	objBuilder.Append "ORDER BY emtt.emp_no, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team ASC"

	Set rsConf = Server.CreateObject("ADODB.RecordSet")
	rsConf.Open objBuilder.ToString(), Dbconn, 1
	objBuilder.Clear()
end if
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html lang="ko">
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
		</script>
		<script type="text/javascript">
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
            function s_sinchung(val, val2, val3, val4, val5) {

            if (!confirm("���������� ��û�Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;

            if (document.getElementById(val3).value == "")
            { alert("��û �뵵�� �������ּ���!"); return; }

			if (document.getElementById(val4).value == "")
            { alert("���ó�� �Է��Ͻʽÿ�!"); return; }

            document.frm.cfm_use.value = document.getElementById(val3).value;
            document.frm.action = "/insa/insa_certificate_print.asp";
            document.frm.submit();
            }
			function s_sinchung2(val, val2, val3, val4, val5) {

            if (!confirm("��������� ��û�Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;

            if (document.getElementById(val3).value == "")
            { alert("��û �뵵�� �������ּ���!"); return; }

			if (document.getElementById(val4).value == "")
            { alert("���ó�� �Է��Ͻʽÿ�!"); return; }

            document.frm.cfm_use.value = document.getElementById(val3).value;
            document.frm.action = "/insa/insa_certificate_career.asp";
            document.frm.submit();
            }

		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_welfare_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_confirm_career_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">���
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">����
                                </label>
							<strong>���� : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�">&nbsp;�����Է��� �˻���ư�� �� Ŭ���Ͻʽÿ�!</a>
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
							<col width="10%" >
							<col width="10%" >
							<col width="12%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="6%" >
                            <col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
								<th scope="col">������</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col" style="background:#FFC">�뵵</th>
								<th scope="col" style="background:#FFC">���ó</th>
                                <th scope="col" style="background:#FFC">���</th>
                                <th colspan="2" scope="col" style="background:#FFC">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, page, page_cnt

						if  view_condi <> "" then
						    i = 0

							do until rsConf.EOF
								i = i + 1
	           			%>
							<tr>
								<td class="first"><%=rsConf("emp_no")%>&nbsp;
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsConf("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rsConf("emp_name")%></a>
                                </td>
                                <td><%=rsConf("emp_job")%>&nbsp;</td>
                                <td><%=rsConf("emp_position")%>&nbsp;</td>
                                <td><%=rsConf("emp_in_date")%>&nbsp;</td>
                                <td><%=rsConf("emp_end_date")%>&nbsp;</td>
                                <td><%=rsConf("emp_company")%>&nbsp;</td>
                                <td><%=rsConf("emp_org_name")%>&nbsp;</td>
                                <td class="left">
                                <select name="cfm_use" id="cfm_use" value="<%=cfm_use%>" style="width:130px">
			            	        <option value="" <% if cfm_use = "" then %>selected<% end if %>>����</option>
				                    <option value='�����' <%If cfm_use = "�����" then %>selected<% end if %>>�����</option>
                                    <option value='������' <%If cfm_use = "������" then %>selected<% end if %>>������</option>
                                    <option value='�б������' <%If cfm_use = "�б������" then %>selected<% end if %>>�б������</option>
                                    <option value='�����������' <%If cfm_use = "�����������" then %>selected<% end if %>>�����������</option>
                                    <option value='���������' <%If cfm_use = "���������" then %>selected<% end if %>>���������</option>
                                    <option value='ȸ�������' <%If cfm_use = "ȸ�������" then %>selected<% end if %>>ȸ�������</option>
                                    <option value='����������' <%If cfm_use = "����������" then %>selected<% end if %>>����������</option>
                                    <option value='���ǻ������' <%If cfm_use = "���ǻ������" then %>selected<% end if %>>���ǻ������</option>
                                    <option value='���ڹ߱޿�' <%If cfm_use = "���ڹ߱޿�" then %>selected<% end if %>>���ڹ߱޿�</option>
                                    <option value='�����' <%If cfm_use = "�����" then %>selected<% end if %>>�����</option>
                                    <option value='�뵿��(û)�����' <%If cfm_use = "�뵿��(û)�����" then %>selected<% end if %>>�뵿��(û)�����</option>
                                    <option value='ī��������' <%If cfm_use = "ī��������" then %>selected<% end if %>>ī��������</option>
                                    <option value='���Ӱ���Ȯ�ο�' <%If cfm_use = "���Ӱ���Ȯ�ο�" then %>selected<% end if %>>���Ӱ���Ȯ�ο�</option>
                                    <option value='���������' <%If cfm_use = "���������" then %>selected<% end if %>>���������</option>
                                    <option value='��ȸ�����' <%If cfm_use = "��ȸ�����" then %>selected<% end if %>>��ȸ�����</option>
                                    <option value='���Ȯ�ο�' <%If cfm_use = "���Ȯ�ο�" then %>selected<% end if %>>���Ȯ�ο�</option>
                                    <option value='������' <%If cfm_use = "������" then %>selected<% end if %>>������</option>
                                </select>
                                </td>
                                <td class="left">
								<input name="cfm_use_dept" type="text" id="cfm_use_dept" style="width:100px" onKeyUp="checklength(this,30)" value="<%=cfm_use_dept%>">
                                </td>
                                <td class="left">
								<input name="cfm_comment" type="text" id="cfm_comment" style="width:100px" onKeyUp="checklength(this,30)" value="<%=cfm_comment%>">
                                </td>
                                <td colspan="2">
                                <input type="image" name="rptCert$ctl01$btnRequest" id="rptCert_ctl01_btnRequest" src="/image/b_certifi.jpg" alt="������� ��û" onclick="s_sinchung2('<%=rsConf("emp_no")%>','<%=rsConf("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
                                </td>
							</tr>
						<%
							rsConf.movenext()
						loop
						rsConf.close() : Set rsConf = Nothing

						end if
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
               <% if user_id = "900002" Or user_id = "102592" then %>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('/insa/insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">������ �߱�</a>
			   <% end if %>
					</div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
                  <input type="hidden" name="in_empno" value="<%=emp_no%>" ID="Hidden1">
                  <input type="hidden" name="in_name" value="<%=emp_name%>" ID="Hidden1">
			</form>
		</div>
	</div>
	</body>
</html>

