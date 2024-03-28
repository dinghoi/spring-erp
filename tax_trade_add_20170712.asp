<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
trade_code = request("trade_code")

trade_no1 = ""
trade_no2 = ""
trade_no3 = ""
trade_name = ""
bill_trade_code = ""
bill_trade_name = ""
trade_id = "일반"
sales_type = ""
trade_owner = ""
trade_addr = ""
trade_uptae = ""
trade_upjong = ""
trade_tel = ""
trade_fax = ""
trade_email = ""
trade_person = ""
trade_person_tel = ""
group_name = ""
use_sw = "Y"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "거래처 등록"
approve_no = request("approve_no")

Sql="select * from tax_bill where approve_no = '"&approve_no&"'"
Set rs=DbConn.Execute(Sql)

trade_no1 = mid(rs("trade_no"),1,3)
trade_no2 = mid(rs("trade_no"),4,2)
trade_no3 = mid(rs("trade_no"),6)
trade_name = rs("trade_name")
trade_id = ""
sales_type = ""
trade_owner = rs("trade_owner")
trade_addr = ""
trade_uptae = ""
trade_upjong = ""
trade_tel = ""
trade_fax = ""
if rs("bill_id") = "1" then	
	person_email = rs("send_email")
  else
	person_email = rs("receive_email")
end if
'	trade_person = rs("trade_person")
'	trade_person_tel = rs("trade_person_tel")
'	bill_trade_code = rs("bill_trade_code")
'	bill_trade_name = rs("bill_trade_name")
rs.close()

sales_saupbu = "Y"
Sql="select * from sales_org where saupbu = '"&saupbu&"'"
Set rs=DbConn.Execute(Sql)
if rs.eof or rs.bof then
	sales_saupbu = "N"
	saupbu = ""
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//ENrs("customer_no")http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.trade_no1.value =="") {
					alert('사업자번호를 입력하세요');
					frm.trade_no1.focus();
					return false;}
				if(document.frm.trade_no2.value =="") {
					alert('사업자번호를 입력하세요');
					frm.trade_no2.focus();
					return false;}
				if(document.frm.trade_no3.value =="") {
					alert('사업자번호를 입력하세요');
					frm.trade_no3.focus();
					return false;}
				if(document.frm.trade_name.value =="") {
					alert('상호를 입력하세요');
					frm.trade_name.focus();
					return false;}
				if(document.frm.sales_type.value =="") {
					alert('거래처 유형을 선택하세요');
					frm.sales_type.focus();
					return false;}
				k = 0;
				for (j=0;j<3;j++) {
					if (eval("document.frm.trade_id[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("계약내용을 선택하시기 바랍니다");
					return false;
				}	
				if(document.frm.trade_owner.value =="") {
					alert('대표자명을 입력하세요');
					frm.trade_owner.focus();
					return false;}
				if(document.frm.trade_addr.value =="") {
					alert('주소를 입력하세요');
					frm.trade_addr.focus();
					return false;}
				if(document.frm.trade_uptae.value =="") {
					alert('업태를 입력하세요');
					frm.trade_uptae.focus();
					return false;}
				if(document.frm.trade_upjong.value =="") {
					alert('업종을 입력하세요');
					frm.trade_upjong.focus();
					return false;}
				if(document.frm.person_email.value !="") {
					if(document.frm.person_name.value =="") {
						alert('계산서메일이 있어 담당자를 입력해야 합니다');
						frm.person_name.focus();
						return false;}}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_trade_add_save.asp" method="post" name="frm">
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
				        <th class="first">사업자번호</th>
				        <td class="left">
                        <input name="trade_no1" type="text" id="trade_no1" style="width:25px; text-align:center" maxlength="3" value="<%=trade_no1%>" onKeyUp="checkNum(this);">
                        -
                        <input name="trade_no2" type="text" id="trade_no2" style="width:20px; text-align:center" maxlength="2" value="<%=trade_no2%>" onKeyUp="checkNum(this);">
                        -
                        <input name="trade_no3" type="text" id="trade_no3" style="width:50px; text-align:center" maxlength="5" value="<%=trade_no3%>" onKeyUp="checkNum(this);"></td>
				        <th>상호</th>
				        <td class="left"><input name="trade_name" type="text" id="trade_name" style="width:200px;" value="<%=trade_name%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">거래처유형</th>
				        <td class="left"><select name="sales_type" id="sales_type" style="width:200px">
				          <option value="">선택</option>
				          <option value="매출" <% if sales_type = "매출" then %>selected<% end if %>>매출</option>
				          <option value="외주" <% if sales_type = "외주" then %>selected<% end if %>>외주</option>
				          <option value="공용" <% if sales_type = "공용" then %>selected<% end if %>>공용</option>
			            </select></td>
				        <th>계약내용</th>
				        <td class="left">
                        <input type="radio" name="trade_id" value="매출" <% if trade_id = "매출" then %>checked<% end if %> style="width:20px">
유지보수
  						<input type="radio" name="trade_id" value="일반" <% if trade_id = "일반" then %>checked<% end if %> style="width:20px">
일반계약
						<input type="radio" name="trade_id" value="계열사" <% if trade_id = "계열사" then %>checked<% end if %> style="width:20px">
Kwon자회사</td>
			          </tr>
				      <tr>
				        <th class="first">그룹명</th>
				        <td class="left"><input name="group_name" type="text" id="group_name" style="width:170px;" value="<%=group_name%>" onKeyUp="checklength(this,30);"><a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="5"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a></td>
				        <th>대표자</th>
				        <td class="left"><input name="trade_owner" type="text" id="trade_owner" style="width:200px;" value="<%=trade_owner%>" onKeyUp="checklength(this,20);"></td>
			          </tr>
				      <tr>
				        <th class="first">주소</th>
				        <td colspan="3" class="left"><input name="trade_addr" type="text" id="trade_addr" style="width:500px" value="<%=trade_addr%>" onKeyUp="checklength(this,100);"></td>
			          </tr>
				      <tr>
				        <th class="first">업태</th>
				        <td class="left"><input name="trade_uptae" type="text" id="trade_uptae" style="width:200px;" value="<%=trade_uptae%>" onKeyUp="checklength(this,50);"></td>
				        <th>업종</th>
				        <td class="left"><input name="trade_upjong" type="text" id="trade_upjong" style="width:200px;" value="<%=trade_upjong%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">전화번호</th>
				        <td class="left"><input name="trade_tel" type="text" id="trade_tel" style="width:200px;" value="<%=trade_tel%>" onKeyUp="checklength(this,20);"></td>
				        <th>팩스</th>
				        <td class="left"><input name="trade_fax" type="text" id="trade_fax" style="width:200px;" value="<%=trade_fax%>" onKeyUp="checklength(this,20);"></td>
			          </tr>
				      <tr>
				        <th class="first">담당자</th>
				        <td class="left"><input name="person_name" type="text" id="person_name" style="width:200px;" value="<%=person_name%>" onKeyUp="checklength(this,20);"></td>
				        <th>담당자 직급</th>
				        <td class="left"><input name="person_grade" type="text" id="person_grade" style="width:200px;" value="<%=person_grade%>" onKeyUp="checklength(this,20);"></td>
			          </tr>
				      <tr>
				        <th class="first">전화번호</th>
				        <td class="left"><input name="person_tel_no" type="text" id="person_tel_no" style="width:200px;" value="<%=person_tel_no%>" onKeyUp="checklength(this,20);"></td>
				        <th>계산서메일</th>
				        <td class="left"><input name="person_email" type="text" id="person_email" style="width:200px;" value="<%=person_email%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">거래처메모</th>
				        <td colspan="3" class="left"><input name="person_memo" type="text" id="person_memo" style="width:500px" value="<%=person_memo%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">케이원담당자</th>
				        <td class="left">
                        <input name="emp_no" type="text" id="emp_no" style="width:80px;" value="<%=emp_no%>" readonly="true">
				        <input name="emp_name" type="text" id="emp_name" style="width:100px;" value="<%=user_name%>" readonly="true">
                        </td>
				        <th>담당사업부</th>
				        <td class="left">
					<% if sales_saupbu = "Y"	then	%>
				        <input name="saupbu" type="text" id="saupbu" style="width:150px;" value="<%=saupbu%>" readonly="true">
					<%   else	%>
                        <select name="saupbu" id="saupbu" style="width:150px">
	                        <option value="" <% if saupbu = "" then %>selected<% end if %>>공통</option>
                    <%
                       Sql="select saupbu from sales_org order by sort_seq asc"
                       rs_org.Open Sql, Dbconn, 1
                       do until rs_org.eof
                    %>
                    		<option value='<%=rs_org("saupbu")%>' <%If saupbu = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                    <%
                    		rs_org.movenext()
                       loop
                       rs_org.close()						
                   	%>
                    	</select>
                    <% end if	%>
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onClick="javascript:frmcheck();" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onClick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="trade_code" value="<%=trade_code%>" ID="Hidden1">
				<input type="hidden" name="bill_trade_code" value="<%=bill_trade_code%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

