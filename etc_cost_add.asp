<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

slip_gubun = "자재"
customer = ""
customer_no = ""
emp_company = ""
org_name = ""
company = ""
account = ""
account1 = ""
price = 0
cost = 0
cost_vat = 0
slip_memo = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "자재 및 장비 비용 등록"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	slip_gubun = rs("slip_gubun")
	customer = rs("customer")
	customer_no = rs("customer_no")
	emp_company = rs("emp_company")
	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
	org_name = rs("org_name")
	company = rs("company")
	account = rs("account")
	account1 = rs("account")
	price = rs("price")
	cost = rs("cost")
	cost_vat = rs("cost_vat")
	slip_memo = rs("slip_memo")
	reg_id = rs("reg_id")
	rs.close()

	title_line = "자재 및 장비 비용 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=slip_date%>" );
			});	  
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
				if(document.frm.slip_date.value <= document.frm.end_date.value) {
					alert('발생일자가 마감이 되어 있는 날자입니다');
					frm.slip_date.focus();
					return false;}
				if(document.frm.slip_date.value > document.frm.curr_date.value) {
					alert('발생일자가 현재일보다 클수가 없습니다.');
					frm.slip_date.focus();
					return false;}
				if(document.frm.end_yn.value =="Y") {
					alert('마감되어 수정 할 수 없습니다');
					frm.end_yn.focus();
					return false;}
				if(document.frm.slip_date.value =="") {
					alert('발생일자를 입력하세요');
					frm.slip_date.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('거래처 선택하세요');
					frm.customer.focus();
					return false;}
				if(document.frm.org_name.value =="") {
					alert('사용조직을 선택하세요');
					frm.org_name.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('비용회사를 선택하세요');
					frm.company.focus();
					return false;}
				if(document.frm.price.value =="") {
					alert('합계금액을 입력하세요');
					frm.price.focus();
					return false;}
				if(document.frm.cost_vat.value =="") {
					alert('부가세를 입력하세요');
					frm.cost_vat.focus();
					return false;}
				if(document.frm.account.value =="" && document.frm.account1.value =="") {
					alert('비용유형을 선택하세요');
					return false;}
				if(document.frm.slip_memo.value =="") {
					alert('발행내역을 입력하세요');
					frm.slip_memo.focus();
					return false;}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function cost_cal(txtObj){
				price = parseInt(document.frm.price.value.replace(/,/g,""));			
				cost_vat = parseInt(document.frm.cost_vat.value.replace(/,/g,""));			
				cost = price - cost_vat;
				cost = String(cost);
				num_len = cost.length;
				sil_len = num_len;
				cost = String(cost);
				if (cost.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) cost = cost.substr(0,num_len -3) + "," + cost.substr(num_len -3,3);
				if (sil_len > 6) cost = cost.substr(0,num_len -6) + "," + cost.substr(num_len -6,3) + "," + cost.substr(num_len -2,3);

				document.frm.cost.value = cost; 

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
					temp=new String(num);
					if(temp.length<1) return "";
					
					// 음수처리
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";
					
					// 소수점이하처리
					dpoint=temp.search(/\./);
					
					if(dpoint>0)
					{
					// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";
					
					// 숫자이외문자 삭제
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);
					
					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);
					
					if(temp.length<4) return minus+temp+dpointVa;
					buf="";
					while (true)
					{
					if(temp.length<3) { buf=temp+buf; break; }
				
					buf=","+temp.substr(temp.length-3)+buf;
					temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);
				
					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";					
			}
			function condi_view() {

				if (eval("document.frm.slip_gubun[0].checked")) {
					document.getElementById('account').style.display = '';
					document.getElementById('account1').style.display = 'none';
				}	
				if (eval("document.frm.slip_gubun[1].checked")) {
					document.getElementById('account1').style.display = '';
					document.getElementById('account').style.display = 'none';
				}	
			}
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "etc_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onLoad="condi_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="etc_cost_add_save.asp" method="post" name="frm">
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
				        <th class="first">발행일자</th>
				        <td class="left">
                        <input name="slip_date" type="text" value="<%=slip_date%>" style="width:80px;text-align:center" id="datepicker">
				          마감일 : <%=end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>외주업체</th>
				        <td class="left">
                        <input name="customer" type="text" value="<%=customer%>" readonly="true" style="width:150px">
                        <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="3"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">사용조직</th>
				        <td class="left">
                          <a href="#" onClick="pop_Window('org_search.asp','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조직조회</a>
                          <input name="org_name" type="text" value="<%=org_name%>" readonly="true" style="width:150px">
                          <input name="emp_company" type="hidden" value="<%=emp_company%>">
                          <input name="bonbu" type="hidden" value="<%=bonbu%>">
                          <input name="saupbu" type="hidden" value="<%=saupbu%>">
                          <input name="team" type="hidden" value="<%=team%>">
                          <input name="reside_place" type="hidden" value="<%=reside_place%>">
                          <input name="reside_company" type="hidden" value="<%=reside_company%>">
                        </td>
				        <th>고객사</th>
				        <td class="left">
                        <input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
                        <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">합계금액</th>
				        <td class="left">
					<% if u_type = "U" then	%>
                        <input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="cost_cal(this);" >
					<%   else	%>
                        <input name="price" type="text" id="price" style="width:100px;text-align:right" onKeyUp="cost_cal(this);" >
                    <% end if	%>
                        </td>
				        <th><span class="first">부가세</span></th>
				        <td class="left">
					<% if u_type = "U" then	%>
                        <input name="cost_vat" type="text" id="cost_vat" style="width:100px;text-align:right" value="<%=formatnumber(cost_vat,0)%>" onKeyUp="cost_cal(this);" >
					<%   else	%>
                        <input name="cost_vat" type="text" id="cost_vat" style="width:100px;text-align:right" onKeyUp="cost_cal(this);" >
                    <% end if	%>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">공급가액</th>
				        <td class="left"><input name="cost" type="text" id="cost" style="width:100px;text-align:right" value="<%=formatnumber(cost,0)%>" readonly="true" ></td>
				        <th>비용구분</th>
				        <td class="left">
                        <input type="radio" name="slip_gubun" value="자재" <% if slip_gubun = "자재" then %>checked<% end if %> style="width:30px" ID="Radio1" onClick="condi_view()">자재
                        <input type="radio" name="slip_gubun" value="장비" <% if slip_gubun = "장비" then %>checked<% end if %> style="width:30px" ID="Radio2" onClick="condi_view()">장비
                        </td>
			          </tr>
				      <tr>
				        <th class="first">비용유형</th>
				        <td class="left">
                        <select name="account" id="account" style="width:150px">
				          <option value="">선택</option>
                          <option value="AS용"  <%If account = "AS용" then %>selected<% end if %>>AS용</option>
                          <option value="랜공사용"  <%If account = "랜공사용" then %>selected<% end if %>>랜공사용</option>
                          <option value="소모성자재"  <%If account = "소모성자재" then %>selected<% end if %>>소모성자재</option>
			            </select>
                        <select name="account1" id="account1" style="width:150px ; display:none">
				          <option value="">선택</option>
                          <option value="상품"  <%If account1 = "상품" then %>selected<% end if %>>상품</option>
                          <option value="백업장비"  <%If account1 = "백업장비" then %>selected<% end if %>>백업장비</option>
                          <option value="집기비품"  <%If account1 = "집기비품" then %>selected<% end if %>>집기비품</option>
                          <option value="공구"  <%If account1 = "공구" then %>selected<% end if %>>공구</option>
			            </select>
                        </td>
				        <th><span class="first">발행내역</span></th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
				<%	
					if u_type = "U" and user_id = reg_id then
						if end_yn = "N" or end_yn = "C" then	
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
						end if
					end if	
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="old_date" value="<%=slip_date%>" ID="Hidden1">
				<input type="hidden" name="customer_no" value="<%=customer_no%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

