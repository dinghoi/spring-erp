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
Dim approve_no, title_line, rs, cost_year, rs_org

approve_no = Request("approve_no")

title_line = "매출 사업부 수정"

objBuilder.Append "SELECT sslt.saupbu, sslt.cost_amt, sslt.vat_amt, sslt.sales_amt, sslt.sales_memo, "
objBuilder.Append "	sslt.sales_date, sslt.sales_company, sslt.emp_name, sslt.emp_no, "
objBuilder.Append "	eomt.org_company, org_name, org_bonbu "
objBuilder.Append "FROM saupbu_sales AS sslt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON sslt.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

cost_year = Mid(rs("sales_date"), 1, 4)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function goAction(){
			   window.close ();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.saupbu.value ==""){
					alert('영업본부를 선택하세요');
					frm.saupbu.focus();
					return false;
				}

				{
					a=confirm('저장 하시겠습니까?')
					if(a==true){
						return true;
					}
					return false;
				}
			}

			//매출 삭제[허정호_20210716]
			function sales_del(grade){
				//권한 체크([memb]account_grade:0 만 가능)
				if(grade !== "0"){
					non_grade();
					return false;
				}

				cfm = confirm("정말 삭제하시겠습니까?");

				if(cfm === true){
					sales_del_Init();
					return;
				}
			}

			function sales_del_Init(){
				var frm = document.frm;
				var app_no = $('#approve_no').val();

				frm.action = "/sales/sales_saupbu_del.asp?approve_no="+app_no;
				frm.submit();
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/sales/sales_saupbu_mod_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">매출일자</th>
				        <td class="left"><%=rs("sales_date")%></td>
				        <th>매출회사</th>
				        <td class="left">
							<%=rs("sales_company")%>
							<%'=rs("org_company")%>
						</td>
			          </tr>
				      <tr>
				        <th class="first">영업본부</th>
				        <td class="left">
	                        <select name="saupbu" id="saupbu" style="width:150px">
                                <!--<option value="회사간거래" <% if rs("saupbu") ="회사간거래" then %>selected<% end if %>>회사간거래</option>-->
							<%
							objBuilder.Append "SELECT saupbu FROM sales_org "
							objBuilder.Append "WHERE sales_year = '"&cost_year&"' AND saupbu <> '기타사업부' "
							objBuilder.Append "ORDER BY sort_seq "

							Set rs_org = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rs_org.EOF
							%>
                                <option value='<%=rs_org("saupbu")%>'
									<%If rs("saupbu") = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%>
								</option>
							<%
								rs_org.MoveNext()
							Loop
							rs_org.Close() : Set rs_org = Nothing

							'박정신 재무이사, 시스템 관리자만 노출
							If user_id = "100359" Or user_id = "102592" Then
							%>
							<option value="기타사업부" <%If rs("org_bonbu") = "기타사업부" Then %>selected<%End If %>>기타사업부</option>
							<%End If%>
	                        </select>
                        </td>
				        <th>담당자</th>
				        <td class="left">
                        	<input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=rs("emp_name")%>" readonly="true">
                        	<input name="emp_no" type="text" id="emp_no" style="width:60px" value="<%=rs("emp_no")%>" readonly="true">
                          	<input name="emp_grade" type="hidden" id="emp_grade" style="width:60px" readonly="true">
							<a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=1','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">공급가액</th>
				        <td class="left"><%=FormatNumber(rs("cost_amt"), 0)%></td>
				        <th>세액</th>
				        <td class="left"><%=FormatNumber(rs("vat_amt"), 0)%></td>
			          </tr>
				      <tr>
				        <th class="first">합계금액</th>
				        <td class="left"><%=FormatNumber(rs("sales_amt"), 0)%></td>
				        <th>품목명</th>
				        <td class="left"><%=rs("sales_memo")%></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01">
						<input type="button" value="변경" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1">
					</span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>&nbsp;&nbsp;
				<%
				If sales_grade = "0" Or empProfitViewAll = "Y" Then
				%>
					<span class="btnType01"><input type="button" value="삭제" onclick="sales_del('<%=account_grade%>');"></span>
				<%
				End If
				%>
                </div>
				<input type="hidden" name="sales_date" value="<%=rs("sales_date")%>" />
				<input type="hidden" name="approve_no" value="<%=approve_no%>" id="approve_no" />
			</form>
			<%
			rs.Close() : Set rs = Nothing
			DBConn.Close() : Set DBConn = Nothing
			%>
		</div>
	</body>
</html>

