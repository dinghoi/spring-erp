<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

customer = ""
customer_no = ""
emp_company = ""
org_name = ""
company = ""
emp_no = user_id
emp_name = user_name
emp_grade = user_grade
account = ""
price = 0
cost = 0
cost_vat = 0
slip_memo = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "외주비 등록"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	customer = rs("customer")
	customer_no = rs("customer_no")
	emp_company = rs("emp_company")
	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
	org_name = rs("org_name")
	company = rs("company")
	account = rs("account")
	price = rs("price")
	cost = rs("cost")
	cost_vat = rs("cost_vat")
	slip_memo = rs("slip_memo")
	emp_no = rs("emp_no")
	emp_name = rs("emp_name")
	emp_grade = rs("emp_grade")
	reg_id = rs("reg_id")
	rs.close()

	title_line = "외주비 변경"
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
					alert('공급자 선택하세요');
					frm.customer.focus();
					return false;}
				if(document.frm.emp_company.value =="") {
					alert('공급받는 회사를 선택하세요');
					frm.emp_company.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('고객사를 선택하세요');
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
				if(document.frm.account.value =="") {
					alert('외주유형을 선택하세요');
					frm.account.focus();
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
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "outside_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="outside_cost_add_save.asp" method="post" name="frm">
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
				        <th>공급자</th>
				        <td class="left">
                        <input name="customer" type="text" value="<%=customer%>" readonly="true" style="width:150px">
                        <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="3"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">공급받는회사</th>
				        <td class="left">
                        <select name="emp_company" id="emp_company" style="width:120px">
				          	<option value="" <% if emp_company = "" then %>selected<% end if %>>선택</option>
                              <%
																' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
																Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                %>
                            <option value='<%=rs_org("org_name")%>' <%If emp_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                              <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()						
                                %>
                        </select>
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
				        <th>외주유형</th>
				        <td class="left">
						<select name="account" id="account" style="width:150px">
				          <option value="" <% if account = "" then %>selected<% end if %>>선택</option>
				          <%
                            Sql="select * from etc_code where etc_type = '43' order by etc_name asc"
                            rs_etc.Open Sql, Dbconn, 1
                            do until rs_etc.eof
							%>
				          <option value='<%=rs_etc("etc_name")%>' <%If account = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
				          <%
                            	rs_etc.movenext()
                            loop
                            rs_etc.close()						
                            %>
			            </select>                        
                        </td>
			          </tr>
				      <tr>
				        <th class="first">담당자</th>
				        <td class="left"><input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>" readonly="true">
                          <input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>" readonly="true">
                        <a href="#" onClick="pop_Window('emp_search.asp','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a></td>
				        <th class="first">발행내역</th>
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
					if u_type = "U" and (user_id = reg_id or user_id = emp_no or cost_grade = "0") then
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
				<input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
				<input type="hidden" name="org_name" value="<%=org_name%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

