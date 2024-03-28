<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
draft_no = request("draft_no")
view_condi = request("view_condi")


draft_man = ""
draft_date = ""
draft_live_id = ""
draft_live_name = ""
draft_tax_id = ""
company = ""
bonbu = ""
saupbu = ""
team = ""
org_name = ""
cost_company = ""
sign_no = ""
deposit_date = ""
deposit_man = ""
work_meno = ""
bank_code = ""
bank_name = ""
account_no = ""
account_name = ""
person_no1 = ""
person_no2 = ""
nation_id = ""
nation_name = ""
tel_ddd = ""
tel_no1 = ""
tel_no2 = ""
hp_ddd = ""
hp_no1 = ""
hp_no2 = ""
e_mail = ""
end_yn = "N"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "사업소득자 등록"

if u_type = "U" then

	sql = "select * from emp_alba_mst where draft_no = '" + draft_no + "'"
	set rs = dbconn.execute(sql)

    draft_no = rs("draft_no")
    draft_man = rs("draft_man")
    draft_date = rs("draft_date")
    draft_live_id = rs("draft_live_id")
    draft_live_name = rs("draft_live_name")
    draft_tax_id = rs("draft_tax_id")
    company = rs("company")
    bonbu = rs("bonbu")
    saupbu = rs("saupbu")
    team = rs("team")
    org_name = rs("org_name")
    cost_company = rs("cost_company")
    sign_no = rs("sign_no")
    deposit_date = rs("deposit_date")
    deposit_man = rs("deposit_man")
    work_memo = rs("work_memo")
    bank_code = rs("bank_code")
    bank_name = rs("bank_name")
    account_no = rs("account_no")
    account_name = rs("account_name")
    person_no1 = rs("person_no1")
    person_no2 = rs("person_no2")
    nation_id = rs("nation_id")
    nation_name = rs("nation_name")
    tel_ddd = rs("tel_ddd")
    tel_no1 = rs("tel_no1")
    tel_no2 = rs("tel_no2")
    hp_ddd = rs("hp_ddd")
    hp_no1 = rs("hp_no1")
    hp_no2 = rs("hp_no2")
    e_mail = rs("e_mail")
    end_yn = rs("end_yn")
	zip_code = rs("zip_code")
    sido = rs("sido")
    gugun = rs("gugun")
    dong = rs("dong")
    addr = rs("addr")
	rs.close()

	title_line = "사업소득자 변경"
end if

    sql="select max(draft_no) as max_seq from emp_alba_mst where company = '"&view_condi&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "800001"
	  else
		max_seq = "000000" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,6)
	end if
    rs_max.close()
	
	if u_type = "U" then
	   code_last = draft_no
	end if
	
draft_no = code_last

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=draft_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=deposit_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
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
				if(document.frm.draft_man.value =="" ) {
					alert('소득자명을 입력하세요');
					frm.draft_man.focus();
					return false;}
				if(document.frm.draft_date.value =="") {
					alert('업무등록일을 입력하세요');
					frm.draft_date.focus();
					return false;}
				if(document.frm.draft_tax_id.value =="") {
					alert('소득구분을 입력하세요');
					frm.draft_tax_id.focus();
					return false;}			
//				if(document.frm.org_name.value =="") {
//					alert('소속을 선택하세요');
//					frm.org_name.focus();
//					return false;}			
				if(document.frm.person_no1.value =="") {
					alert('주민등록번호를 입력하세요');
					frm.person_no1.focus();
					return false;}			
				if(document.frm.account_no.value =="" ) {
					alert('계좌번호를 입력 하세요');
					frm.account_no.focus();
					return false;}
				if(document.frm.account_name.value =="" ) {
					alert('예금주명을 입력 하세요');
					frm.account_name.focus();
					return false;}
				if(document.frm.bank_name.value =="" ) {
					alert('은행을 선택 하세요');
					frm.bank_name.focus();
					return false;}
			
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
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
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_alba_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
						    <col width="22%" >
						    <col width="10%" >
						    <col width="22%" >
						    <col width="10%" >
						    <col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">등록번호</th>
                                <td class="left"><%=draft_no%><input name="draft_no" type="hidden" value="<%=draft_no%>"></td>
								<th>소득자명</th>
								<td class="left">
                                <input name="draft_man" type="text" value="<%=draft_man%>" style="width:120px" onKeyUp="checklength(this,20)"></td>
                                <th>업무등록일</th>
								<td class="left"><input name="draft_date" type="text" value="<%=draft_date%>" style="width:70px" id="datepicker"></td>
							</tr>
 							<tr>
								<th class="first">거주구분</th>
								<td class="left">
                                <select name="draft_live_name" id="draft_live_name" style="width:120px">
								  <option value="">선택</option>
								  <option value="거주" <%If draft_live_name = "거주" then %>selected<% end if %>>거주</option>
								  <option value="비거주" <%If draft_live_name = "비거주" then %>selected<% end if %>>비거주</option>
							    </select>
                                </td>
								<th>소득구분</th>
								<td class="left"><select name="draft_tax_id" id="draft_tax_id" style="width:120px">
								  <option value="">선택</option>
								  <option value="심부름용역" <%If draft_tax_id = "심부름용역" then %>selected<% end if %>>심부름용역</option>
								  <option value="자문/고문" <%If draft_tax_id = "자문/고문" then %>selected<% end if %>>자문/고문</option>
                                  <option value="작곡가" <%If draft_tax_id = "작곡가" then %>selected<% end if %>>작곡가</option>
							    </select></td>
                                <th>내/외국인</th>
								<td class="left"><select name="nation_name" id="nation_name" style="width:120px">
								  <option value="">선택</option>
								  <option value="내국인" <%If nation_name = "내국인" then %>selected<% end if %>>내국인</option>
								  <option value="외국인" <%If nation_name = "외국인" then %>selected<% end if %>>외국인</option>
							    </select></td>
							</tr>
                            <tr>
								<th class="first">소속</th>
								<td colspan="5" class="left">
                                <input name="org_name" type="text" id="org_name" style="width:100px" value="<%=org_name%>" readonly="true">
                                -
                                <input name="company" type="text" id="company" style="width:100px" value="<%=company%>" readonly="true">
                                <input name="bonbu" type="text" id="bonbu" style="width:100px" value="<%=bonbu%>" readonly="true">
                                <input name="saupbu" type="text" id="saupbu" style="width:100px" value="<%=saupbu%>" readonly="true">
                                <input name="team" type="text" id="team" style="width:100px" value="<%=team%>" readonly="true">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="alba"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=850,height=400')">부서찾기</a>
                                </td>
							</tr>
                           	<tr>
								<th class="first">주민번호</th>
								<td class="left">
                                <input name="person_no1" type="text" value="<%=person_no1%>" style="width:50px" onKeyUp="checklength(this,7)">
                                -
                                <input name="person_no2" type="text" value="<%=person_no2%>" style="width:50px" onKeyUp="checklength(this,8)">
                                </td>
								<th>전화번호</th>
                                <td class="left">
								<input name="tel_ddd" type="text" id="tel_ddd" size="3" maxlength="3" value="<%=tel_ddd%>" >
								  -
                                <input name="tel_no1" type="text" id="tel_no1" size="4" maxlength="4" value="<%=tel_no1%>" >
                                  -
                                <input name="tel_no2" type="text" id="tel_no2" size="4" maxlength="4" value="<%=tel_no2%>" >
                                </td>
                                <th>핸드폰</th>
                                <td class="left">
								<input name="hp_ddd" type="text" id="hp_ddd" size="3" maxlength="3" value="<%=hp_ddd%>" >
								  -
                                <input name="hp_no1" type="text" id="hp_no1" size="4" maxlength="4" value="<%=hp_no1%>" >
                                  -
                                <input name="hp_no2" type="text" id="hp_no2" size="4" maxlength="4" value="<%=hp_no2%>" >
                                </td>
                            <tr>
								<th>주소(현)</th>
								<td colspan="5" class="left">
								<input name="sido" type="text" id="sido" style="width:80px" readonly="true" value="<%=sido%>">
              					<input name="gugun" type="text" id="gugun" style="width:100px" readonly="true" value="<%=gugun%>">
              					<input name="dong" type="text" id="dong" style="width:120px" readonly="true" value="<%=dong%>">
              					<input name="addr" type="text" id="addr" style="width:200px" value="<%=addr%>" >
              					<input name="zip_code" type="hidden" id="zip_code" value="<%=zip_code%>">
              					<a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="alba"%>','family_zip_select','scrollbars=yes,width=600,height=400')">주소조회</a>
                                </td>
                             </tr>
                            <tr>          
                                <th class="first">은행명</th>
                                <td class="left">
					         <%
					            Sql="select * from emp_etc_code where emp_etc_type = '50' order by emp_etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
					         %>
					            <select name="bank_name" id="bank_name" style="width:120px">
                                  <option value="" <% if bank_name = "" then %>selected<% end if %>>선택</option>
                			 <% 
								do until rs_etc.eof 
			  				 %>
                		  		 <option value='<%=rs_etc("emp_etc_name")%>' <%If bank_name = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			 <%
									rs_etc.movenext()  
								loop 
							    rs_etc.Close()
							 %>
            		            </select>                 
                                </td>
                                <th>계좌번호</th>
                                <td class="left">
								<input name="account_no" type="text" id="account_no" value="<%=account_no%>" style="width:120px" onKeyUp="checklength(this,30)">
                                </td>
                                <th>예금주</th>
                                <td class="left">
								<input name="account_name" type="text" id="account_name" value="<%=account_name%>" style="width:120px" onKeyUp="checklength(this,20)">
                                </td>
							</tr>
							<tr>
                                <th class="first">e_mal</th>
                                <td colspan="3" class="left">
								<input name="e_mail" type="text" id="e_mail" value="<%=e_mail%>" style="width:150px" onKeyUp="checklength(this,30)">
                                </td>
                                <th>해지여부</th>
                                <td class="left">
                                <input type="radio" name="end_yn" value="Y" <% if end_yn = "Y" then %>checked<% end if %> style="width:40px" id="Radio1">해지
                                <input type="radio" name="end_yn" value="N" <% if end_yn = "N" then %>checked<% end if %> style="width:40px" id="Radio2">안함
                                </td>
							</tr>

							<tr>
								<th class="first">전자결재<br>No.</th>
								<td class="left">
                                <input name="sign_no" type="text" value="<%=sign_no%>" style="width:120px" onKeyUp="checklength(this,20)"></td>
								<th>기안일자</th>
								<td class="left">
                                <input name="deposit_date" type="text" value="<%=deposit_date%>" style="width:70px" id="datepicker1"></td>
                                <th>기안자</th>
								<td class="left">
                                <input name="deposit_man" type="text" value="<%=deposit_man%>" style="width:120px" onKeyUp="checklength(this,30)"></td>
                                </td>
							</tr>                                                    
							<tr>
								<th class="first">작업내용</th>
								<td colspan="5" class="left">
                                <input name="work_memo" type="text" value="<%=work_memo%>" style="width:550px" onKeyUp="checklength(this,50)"></td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

