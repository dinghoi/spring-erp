<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
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
Dim page_cnt, pg_cnt, Page, be_pg, view_condi
Dim field_bonbu, field_saupbu, field_team, field_org_name
Dim field_org_code, field_reside_company
Dim view_c, pgsize, start_page, stpage
Dim total_record, total_page, title_line
Dim pg_url, searchTxt
Dim strSql, rsOrg, arrOrg


page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
view_condi = f_Request("view_condi")
searchTxt = f_Request("searchTxt")
view_c = f_Request("view_c")

title_line = " ���� ��Ȳ "
be_pg = "/insa/insa_org_mg.asp"


If view_condi = "" Then
	view_condi = "���̿�"
End If

If view_c = "" Then
	view_c = "bonbu"
End If

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

strSql = "CALL USP_INSA_ORG_MST_LIST('"&view_condi&"', '"&view_c&"', '"&searchTxt&"', "&stpage&", "&pgsize&")"

Set rsOrg = DBConn.Execute(strSql)

If Not rsOrg.EOF Then
	arrOrg = rsOrg.getRows()
	total_record = CInt(arrOrg(0, 0))
Else
	total_record = 0
End If

Call Rs_Close(rsOrg)

pg_url = "&view_condi="&view_condi&"&view_c="&view_c&"&searchTxt="&searchTxt
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			$(document).ready(function(){
				condi_view();
			});

			function searchEnter(name){
				$("input[name = "+name+"]").on("keyup", function(e){
					if(e.keyCode === 13){
						frmcheck();
					}
				});
			}

			//���� �޴� �̺�Ʈ
			function getPageCode(){
				return "5 1";
			}

			function frmcheck(){
				//if (formcheck(document.frm) && chkfrm()) {
				if (chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == ""){
					alert("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			function condi_view(){
				if(eval("document.frm.view_c[0].checked")){
					document.getElementById('bonbu1').style.display = '';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('org_name1').style.display = 'none';
					document.getElementById('reside_company1').style.display = 'none';
					document.getElementById('org_code1').style.display = 'none';
				}

				if(eval("document.frm.view_c[1].checked")){
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = '';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('org_name1').style.display = 'none';
					document.getElementById('reside_company1').style.display = 'none';
					document.getElementById('org_code1').style.display = 'none';
				}

				if(eval("document.frm.view_c[2].checked")){
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = '';
					document.getElementById('org_name1').style.display = 'none';
					document.getElementById('reside_company1').style.display = 'none';
					document.getElementById('org_code1').style.display = 'none';
				}

				if(eval("document.frm.view_c[3].checked")){
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('org_name1').style.display = '';
					document.getElementById('reside_company1').style.display = 'none';
					document.getElementById('org_code1').style.display = 'none';
				}

				if(eval("document.frm.view_c[4].checked")){
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('org_name1').style.display = 'none';
					document.getElementById('reside_company1').style.display = '';
					document.getElementById('org_code1').style.display = 'none';
				}

				if(eval("document.frm.view_c[5].checked")){
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
					document.getElementById('org_name1').style.display = 'none';
					document.getElementById('reside_company1').style.display = 'none';
					document.getElementById('org_code1').style.display = '';
				}
			}

			//���� ����[����ȣ_20210729]
			function insaOrgMod(code, condi){
				var url = '/insa/insa_org_reg.asp';
				var pop_name = '���� ����';
				var features = 'scrollbars=yes,width=1250,height=400';
				var param;

				param = '?org_code='+code+'&view_condi='+condi+'&u_type=U';

				url += param;

				pop_Window(url, pop_name, features);
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>�˻� ����</dt>
                        <dd>
                            <p>
                                <strong>ȸ��</strong>
							    <label>
								<%
								'ȸ��� �˻�(selectbox)[����ȣ_20210601]
								'SelectEmpOrgList(name, id, css, ������)
								Call SelectEmpOrgList("view_condi", "view_condi", "width:150px", view_condi)
								%>
                                </label>

								<label>
									<input type="radio" name="view_c" value="bonbu" <%If view_c = "bonbu" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">����
								</label>
								<label>
									<input type="radio" name="view_c" value="saupbu" <%If view_c = "saupbu" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">�����
								</label>
				                <label>
									<input type="radio" name="view_c" value="team" <%If view_c = "team" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">��
								</label>
								<label>
									<input type="radio" name="view_c" value="org_name" <%If view_c = "org_name" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">������
								</label>
								<label>
									<input type="radio" name="view_c" value="reside_company" <%If view_c = "reside_company" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">���� ȸ��
								</label>
								<%
								If SysAdminYn = "Y" Then	'�ý��� ������ ���� ����[����ȣ_20210728]
								%>
								<label>
									<input type="radio" name="view_c" value="org_code" <%If view_c = "org_code" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">�ڵ�
								</label>
								<%End If%>

                                <label id="bonbu1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���θ�</strong>
								</label>
								<label id="saupbu1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����θ�</strong>
								</label>
                                <label id="team1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����</strong>
								</label>
								<label id="org_name1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������</strong>
								</label>
								<label id="reside_company1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����ȸ��</strong>
								</label>
								<label id="org_code1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ڵ�</strong>
								</label>

								<input name="searchTxt" type="text" value="<%=searchTxt%>" style="width:120px; text-align:left; ime-mode:active" id="field_view" onkeypress="searchEnter('searchTxt');">

								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="4%" >
				      <col width="9%" >
                      <col width="6%" >
                      <col width="4%" >
				      <col width="5%" >
				      <col width="6%" >
                      <col width="8%" >
				      <col width="8%" >
					  <col width="8%" >
				      <col width="8%" >
				      <col width="8%" >
                      <col width="11%" >
				      <col width="6%" >
                      <!--<col width="5%" >
				      <col width="5%" >
                      <col width="3%" >-->
			        </colgroup>
				    <thead>
				      <tr>
				        <th colspan="3" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;��&nbsp;&nbsp;��</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
				        <th rowspan="2" scope="col">����ȸ��</th>
						<th rowspan="2" scope="col">����ó</th>
                        <th rowspan="2" scope="col">����������</th>
				        <!--<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">����&nbsp;������</th>-->
                        <th rowspan="2" scope="col">����</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">�ڵ�</th>
				        <th scope="col">������</th>
                        <th scope="col">����<br>����</th>
				        <th scope="col">���</th>
				        <th scope="col">����</th>
                        <th scope="col">ȸ&nbsp;&nbsp;��</th>
				        <th scope="col">��&nbsp;&nbsp;��</th>
						<th scope="col">�����</th>
				        <th scope="col">��</th>
				        <!--<th scope="col">���</th>
                        <th scope="col">����</th>-->
                      </tr>
			        </thead>
				    <tbody>
					<%
					Dim i
					Dim org_code, org_level, org_table_org, org_empno, org_emp_name
					Dim org_company, org_saupbu, org_team, org_reside_company, org_date
					Dim org_owner_empno, org_owner_empname, org_bonbu
					Dim org_reside_place, trade_code

					If IsArray(arrOrg) Then
						For i = LBound(arrOrg) To UBound(arrOrg, 2)
							org_code = arrOrg(1, i)
							org_name = arrOrg(2, i)
							org_level = arrOrg(3, i)
							org_table_org = arrOrg(4, i)
							org_empno = arrOrg(5, i)
							org_emp_name = arrOrg(6, i)
							org_company = arrOrg(7, i)
							org_bonbu = arrOrg(8, i)
							org_saupbu = arrOrg(9, i)
							org_team = arrOrg(10, i)
							org_reside_company = arrOrg(11, i)
							org_date = arrOrg(12, i)
							org_owner_empno = arrOrg(13, i)
							org_owner_empname = arrOrg(14, i)
							org_reside_place = arrOrg(15, i)
							trade_code = arrOrg(16, i)
					%>
				      <tr>
				        <td class="first"><%=org_code%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('/insa/insa_org_view.asp?org_code=<%=org_code%>','insa_org_view_pop','scrollbars=yes,width=750,height=360')"><%=org_name%></a>&nbsp;</td>
                        <td><%=org_level%>&nbsp;</td>
                        <td><%=org_empno%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=org_empno%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=670')"><%=org_emp_name%></a>
						</td>
                        <td><%=org_company%>&nbsp;</td>
				        <td><%=org_bonbu%>&nbsp;</td>
						<td><%=org_saupbu%>&nbsp;</td>
                        <td><%=org_team%>&nbsp;</td>
                        <td><%=org_reside_company%>&nbsp;</td>
						<td><%=org_reside_place%>&nbsp;</td>
                        <td><%=org_date%>&nbsp;</td>
                        <!--<td><%'=org_owner_empno%>&nbsp;</td>
                        <td><%'=org_owner_empname%>&nbsp;</td>-->
                        <td><a href="#" onclick="insaOrgMod('<%=org_code%>', '<%=view_condi%>');">����</a>&nbsp;</td>
			          </tr>
				      <%
					  		Next
						Else
							Response.Write "<tr><td colspan='13' style='text-weight:bold;'>�ش� ������ �����ϴ�.</td></tr>"
						End If
					  %>
			        </tbody>
			      </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="/insa/excel/insa_excel_org.asp?view_condi=<%=view_condi%>&view_c=<%=view_c%>&searchTxt=<%=searchTxt%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)

					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('/insa/insa_org_reg.asp?view_condi=<%=view_condi%>','insa_org_reg_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">�ű��������</a>
					</div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>

		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
        <!--<input type="hidden" name="field_check" value="<%'=field_view%>" ID="field_check">-->
	</body>
</html>