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

title_line = " 조직 현황 "
be_pg = "/insa/insa_org_mg.asp"


If view_condi = "" Then
	view_condi = "케이원"
End If

If view_c = "" Then
	view_c = "bonbu"
End If

pgsize = 10 ' 화면 한 페이지

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
		<title>인사 관리 시스템</title>
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

			//서브 메뉴 이벤트
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
					alert("필드조건을 선택하시기 바랍니다");
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

			//조직 수정[허정호_20210729]
			function insaOrgMod(code, condi){
				var url = '/insa/insa_org_reg.asp';
				var pop_name = '조직 변경';
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
					<legend>조회영역</legend>
					<dl>
						<dt>검색 조건</dt>
                        <dd>
                            <p>
                                <strong>회사</strong>
							    <label>
								<%
								'회사명 검색(selectbox)[허정호_20210601]
								'SelectEmpOrgList(name, id, css, 조직명)
								Call SelectEmpOrgList("view_condi", "view_condi", "width:150px", view_condi)
								%>
                                </label>

								<label>
									<input type="radio" name="view_c" value="bonbu" <%If view_c = "bonbu" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">본부
								</label>
								<label>
									<input type="radio" name="view_c" value="saupbu" <%If view_c = "saupbu" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">사업부
								</label>
				                <label>
									<input type="radio" name="view_c" value="team" <%If view_c = "team" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">팀
								</label>
								<label>
									<input type="radio" name="view_c" value="org_name" <%If view_c = "org_name" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">조직명
								</label>
								<label>
									<input type="radio" name="view_c" value="reside_company" <%If view_c = "reside_company" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">상주 회사
								</label>
								<%
								If SysAdminYn = "Y" Then	'시스템 관리자 권한 여부[허정호_20210728]
								%>
								<label>
									<input type="radio" name="view_c" value="org_code" <%If view_c = "org_code" Then%>checked<%End If%> style="width:25px;" onClick="condi_view();">코드
								</label>
								<%End If%>

                                <label id="bonbu1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;본부명</strong>
								</label>
								<label id="saupbu1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;사업부명</strong>
								</label>
                                <label id="team1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;팀명</strong>
								</label>
								<label id="org_name1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;조직명</strong>
								</label>
								<label id="reside_company1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;상주회사</strong>
								</label>
								<label id="org_code1">
									<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;코드</strong>
								</label>

								<input name="searchTxt" type="text" value="<%=searchTxt%>" style="width:120px; text-align:left; ime-mode:active" id="field_view" onkeypress="searchEnter('searchTxt');">

								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
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
				        <th colspan="3" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;직&nbsp;&nbsp;장</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
				        <th rowspan="2" scope="col">상주회사</th>
						<th rowspan="2" scope="col">상주처</th>
                        <th rowspan="2" scope="col">조직생성일</th>
				        <!--<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">상위&nbsp;조직장</th>-->
                        <th rowspan="2" scope="col">수정</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">코드</th>
				        <th scope="col">조직명</th>
                        <th scope="col">조직<br>구분</th>
				        <th scope="col">사번</th>
				        <th scope="col">성명</th>
                        <th scope="col">회&nbsp;&nbsp;사</th>
				        <th scope="col">본&nbsp;&nbsp;부</th>
						<th scope="col">사업부</th>
				        <th scope="col">팀</th>
				        <!--<th scope="col">사번</th>
                        <th scope="col">성명</th>-->
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
                        <td><a href="#" onclick="insaOrgMod('<%=org_code%>', '<%=view_condi%>');">수정</a>&nbsp;</td>
			          </tr>
				      <%
					  		Next
						Else
							Response.Write "<tr><td colspan='13' style='text-weight:bold;'>해당 내용이 없습니다.</td></tr>"
						End If
					  %>
			        </tbody>
			      </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="/insa/excel/insa_excel_org.asp?view_condi=<%=view_condi%>&view_c=<%=view_c%>&searchTxt=<%=searchTxt%>" class="btnType04">엑셀다운로드</a>
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
                    <a href="#" onClick="pop_Window('/insa/insa_org_reg.asp?view_condi=<%=view_condi%>','insa_org_reg_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">신규조직등록</a>
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