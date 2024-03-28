<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--include virtual="/include/end_check.asp" -->
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
Dim slip_date, slip_seq
Dim end_saupbu, sql, rs_end, new_date, end_date
Dim rsCost
Dim slip_gubun, customer, customer_no
Dim company, account, account_item, price, cost, cost_vat, slip_memo
Dim emp_name, emp_grade, reg_id, mg_saupbu, pl_yn
Dim account_view, title_line
Dim u_type, end_yn

slip_date = f_Request("slip_date")
slip_seq = f_Request("slip_seq")

If saupbu = "" Then
	end_saupbu = "사업부외나머지"
Else
  	end_saupbu = saupbu
End If

sql = "SELECT MAX(end_month) as max_month " &_
      "  FROM cost_end                    " &_
     " WHERE saupbu = '"&end_saupbu&"'   " &_
     "   AND end_yn ='Y'                 "
objBuilder.Append "SELECT MAX(end_month) as max_month "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE saupbu = '"&end_saupbu&"' "
objBuilder.Append "	AND end_yn ='Y'"

Set rs_end = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_end("max_month")) Then
	end_date = "2014-08-31"
Else
	new_date = DateAdd("m", 1, DateValue(Mid(rs_end("max_month"), 1, 4) & "-" & Mid(rs_end("max_month"), 5, 2) & "-01"))
	end_date = DateAdd("d", -1, new_date)
End If

rs_end.Close() : Set rs_end = Nothing

'Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"	'기존 주석
'sql = "SELECT * "
'sql = sql & "FROM general_cost AS gect "
'sql = sql & "INNER JOIN emp_master AS emtt ON gect.emp_no = emtt.emp_no "
'sql = sql & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
'sql = sql & "WHERE slip_date = '"&slip_date&"' AND slip_seq = '"&slip_seq&"' "

objBuilder.Append "SELECT slip_gubun, customer, customer_no, gect.emp_company, bonbu, saupbu, team, "
objBuilder.Append "	gect.org_name, company, account, account_item, price, cost, cost_vat, slip_memo,"
objBuilder.Append "	gect.emp_no, gect.emp_name, gect.emp_grade, reg_id, gect.mg_saupbu, pl_yn, slip_date, "
objBuilder.Append "	org_company, slip_seq, approve_no "
objBuilder.Append "FROM general_cost AS gect "
objBuilder.Append "INNER JOIN emp_master AS emtt ON gect.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE slip_date = '"&slip_date&"' AND slip_seq = '"&slip_seq&"' "

Set rsCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

slip_gubun = rsCost("slip_gubun")
customer = rsCost("customer")
customer_no = rsCost("customer_no")
emp_company = rsCost("emp_company")
bonbu = rsCost("bonbu")
saupbu = rsCost("saupbu")
team = rsCost("team")
org_name = rsCost("org_name")'
company = rsCost("company")
account = rsCost("account")
account_item = rsCost("account_item")
price = rsCost("price")
cost = rsCost("cost")
cost_vat = rsCost("cost_vat")
slip_memo = rsCost("slip_memo")
emp_no = rsCost("emp_no")
emp_name = rsCost("emp_name")
emp_grade = rsCost("emp_grade")
reg_id = rsCost("reg_id")
mg_saupbu = rsCost("mg_saupbu")
pl_yn = rsCost("pl_yn")

If slip_gubun = "비용" Then
	account_view = account & "-" & account_item
Else
  	account_view = account_item
End If

title_line = "매입 세금계산서 수정("&account_view&")"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
			   window.close();
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
//				if(document.frm.slip_date.value <= document.frm.end_date.value) {
//					alert('발행일자가 마감이 되어 있는 날자입니다');
//					frm.slip_date.focus();
//					return false;}
				if(document.frm.mg_saupbu.value =="선택"){
					alert('담당영업사업부를 선택하세요');
					frm.mg_saupbu.focus();
					return false;
				}

				if(document.frm.company.value ==""){
					alert('고객사를 선택하세요');
					frm.company.focus();
					return false;
				}

				if(document.frm.slip_gubun.value ==""){
					alert('비용유형을 선택하세요');
					frm.slip_gubun.focus();
					return false;
				}

				//if(document.frm.company.value =="공통" || document.frm.company.value =="케이원정보통신"){
				if(document.frm.company.value =="공통" || document.frm.company.value =="케이원"){
					if(document.frm.mg_saupbu.value != ""){
						//console.log(document.frm.mg_saupbu.value);
						//console.log(document.frm.saupbu.value);
						//if(document.frm.mg_saupbu.value != document.frm.saupbu.value){
						if(document.frm.mg_saupbu.value != document.frm.bonbu.value){
							alert('고객사가 공통인 경우 사용조직사업부와 담당영업사업부를 동일해야합니다.');
							frm.org_name.focus();
							return false;
						}
					}
				}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true){
					return true;
				}
				return false;
				}
			}

			function pl_view(){
				var d = document.frm.cost_grade.value;

				if(d == '0'){
					document.getElementById('pl_col').style.display = '';
				}
			}

			function delcheck(){
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "tax_bill_in_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
        </script>
	</head>
	<body onload="pl_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/sales/tax_bill_in_mod_save.asp" method="post" name="frm">
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
				        <th class="first">발행일자</th>
				        <td class="left"><%=rsCost("slip_date")%>&nbsp;
				          마감일 : <%=end_date%>
                        </td>
				        <th>공급받는회사</th>
				        <td class="left"><%=rsCost("emp_company")%></td>
			          </tr>
				      <tr>
				        <th class="first">사용조직</th>
				        <td class="left">
                        <input name="org_name" type="text" readonly="true" value="<%=org_name%>" style="width:150px">
                        <%=rsCost("org_company")%><a href="#" onClick="pop_Window('/org_search.asp?gubun=계산서&org_company=<%=rsCost("emp_company")%>','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
				        <input name="emp_company" type="hidden" value="<%=emp_company%>">
				        <input name="bonbu" type="hidden" value="<%=bonbu%>">
				        <input name="saupbu" type="hidden" value="<%=saupbu%>">
				        <input name="team" type="hidden" value="<%=team%>">
				        <input name="reside_place" type="hidden" value="<%=reside_place%>">
                        <input name="reside_company" type="hidden" value="<%=reside_company%>">
                        </td>
				        <th>담당영업사업부</th>
				        <td class="left">
						<%
						Dim cost_year, sql_org, rs_org

						cost_year = Mid(rsCost("slip_date"), 1, 4)

						sql_org = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
                        Set rs_org = DBConn.Execute(sql_org)
                        %>
                          <select name="mg_saupbu" id="mg_saupbu" style="width:150px">
                            <option value='선택' <%If mg_saupbu = "선택" Then %>selected<% End If %>>선택</option>
                            <option value='' <%If mg_saupbu = "" Then %>selected<% End If %>>담당영업부없음</option>
                            <%
                                Do Until rs_org.EOF
                            %>
                            <option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = mg_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                            <%
                                    rs_org.MoveNext()
                                Loop
                                rs_org.Close() : Set rs_org = Nothing
                            %>
                        </select></td>
			          </tr>
				      <tr>
				        <th class="first">공급자</th>
				        <td class="left">
							<%=Mid(rsCost("customer_no"), 1, 3)%> - <%=Mid(rsCost("customer_no"), 4, 2)%> - <%=Right(rsCost("customer_no"), 5)%>&nbsp;<%=rsCost("customer")%>
						</td>
				        <th>담당자</th>
				        <td class="left">
							<input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>" readonly="true">
							<input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>" readonly="true">
							<a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=1','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">발행내역</th>
				        <td class="left">
							<input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=rsCost("slip_memo")%>">
						</td>
				        <th>금액</th>
				        <td class="left">
							<strong>공급가액 :</strong>&nbsp;<%=FormatNumber(rsCost("cost"), 0)%>&nbsp;&nbsp;&nbsp;
							<strong>부가세 :</strong>&nbsp;<%=FormatNumber(rsCost("cost_vat"), 0)%>
						</td>
			          </tr>
				      <tr>
				        <th class="first">고객사</th>
				        <td class="left">
							<input name="company" type="text" value="<%=rsCost("company")%>" readonly="true" style="width:150px">
							<a href="#" onClick="pop_Window('/trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
						</td>
				        <th>비용유형</th>
				        <td class="left">
							<input type="text" name="slip_gubun" ID="slip_gubun" readonly="true" style="width:100px" value="<%=rsCost("slip_gubun")%>">
							<input name="account_view" type="text" readonly="true" style="width:150px" value="<%=account_view%>">
							<a href="#" onClick="pop_Window('/tax_bill_account_search.asp','tax_bill_account_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
							<input name="account" type="hidden" id="account" value="<%=rsCost("account")%>">
							<input name="account_item" type="hidden" id="account_item" value="<%=rsCost("account_item")%>">
						</td>
			          </tr>
				      <tr id="pl_col" style="display:none">
				        <th class="first">손익포함</th>
				        <td colspan="3" class="left">
							<input type="radio" name="pl_yn" value="Y" <%If pl_yn = "Y" Then %>checked<% End If %> style="width:30px" id="Radio2">손익포함
							<input type="radio" name="pl_yn" value="N" <%If pl_yn = "N" Then %>checked<% End If %> style="width:30px" id="Radio">손익미포함
						</td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
				<%'	if end_yn = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%'	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
				<%
					If (user_id = reg_id Or user_id = emp_no) Then
						If end_yn <> "Y" Then
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%
						End If
					End If
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="slip_date" value="<%=rsCost("slip_date")%>" ID="Hidden1">
				<input type="hidden" name="slip_seq" value="<%=rsCost("slip_seq")%>" ID="Hidden1">
				<input type="hidden" name="approve_no" value="<%=rsCost("approve_no")%>" ID="Hidden1">
				<input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="cost_grade" value="<%=cost_grade%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
rsCost.Close() : Set rsCost = Nothing
DBConn.Close() : Set DBConn = Nothing
%>
