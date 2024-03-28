<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

u_type = request("u_type")
overtime_code = request("overtime_code")

sql = "select * from overtime_code order by cost_detail desc"
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
	sql = "select * from overtime_code where overtime_code = '" + overtime_code + "'"
	Set rs_etc=DbConn.Execute(Sql)
	holi_id = rs_etc("holi_id")
	work_gubun = rs_etc("work_gubun")
	apply_dept = rs_etc("apply_dept")
	apply_unit = rs_etc("apply_unit")
	overtime_amt = rs_etc("overtime_amt")
	meals_yn = rs_etc("meals_yn")
	work_time1 = rs_etc("work_time1")
	work_time2 = rs_etc("work_time2")
	sign_yn = rs_etc("sign_yn")
	you_yn = rs_etc("you_yn")
	use_yn = rs_etc("use_yn")
	overtime_memo = rs_etc("overtime_memo")
  else
	holi_id = ""
	work_gubun = ""
	apply_dept = ""
	apply_unit = ""
	overtime_amt = 0
	meals_yn = ""
	work_time1 = ""
	work_time2 = 0
	sign_yn = ""
	you_yn = ""
	use_yn = "Y"
	overtime_memo = ""
end if	

title_line = "야특근 수당 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmsubmit () {
				document.condi_frm.submit ();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.holi_id.value =="") {
					alert('휴일구분을 선택하세요');
					frm.holi_id.focus();
					return false;}
				if(document.frm.work_gubun.value =="") {
					alert('수당명을 입력하세요');
					frm.work_gubun.focus();
					return false;}
				if(document.frm.apply_dept.value =="") {
					alert('적용부서를 입력하세요');
					frm.apply_dept.focus();
					return false;}
				if(document.frm.apply_unit.value =="") {
					alert('기준을 선택하세요');
					frm.apply_unit.focus();
					return false;}
				if(document.frm.overtime_amt.value == 0) {
					alert('단가를 입력하세요');
					frm.overtime_amt.focus();
					return false;}

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.meals_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("식대포함 여부를 선택하세요");
					return false;
				}	

				if(document.frm.work_time1.value =="") {
					alert('근무시간을 입력하세요');
					frm.work_time1.focus();
					return false;}
				
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.sign_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("사전품의 여부를 선택하세요");
					return false;
				}	
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.you_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("유무상 여부를 선택하세요");
					return false;
				}	
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.use_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("사용 여부를 선택하세요");
					return false;
				}	
			
				if(document.frm.overtime_memo.value =="") {
					alert('비고를 입력하세요');
					frm.overtime_memo.focus();
					return false;}

				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
			function frmcancel() 
				{
					document.frm.action = "overtime_code_mg.asp?u_type=''";
					document.frm.submit();
				}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="75%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="8%" >
				          <col width="*" >
				          <col width="7%" >
				          <col width="10%" >
				          <col width="5%" >
				          <col width="7%" >
				          <col width="5%" >
				          <col width="10%" >
				          <col width="5%" >
				          <col width="6%" >
				          <col width="5%" >
				          <col width="16%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">수당구분</th>
				            <th scope="col">수당명</th>
				            <th scope="col">휴일구분</th>
				            <th scope="col">적용부서</th>
				            <th scope="col">기준</th>
				            <th scope="col">단가</th>
				            <th scope="col">식대</th>
				            <th scope="col">근무시간1</th>
				            <th scope="col">사전<br>품의</th>
				            <th scope="col">유무상</th>
				            <th scope="col">사용</th>
				            <th scope="col">비고</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        do until rs.eof
							if rs("you_yn") = "Y" then
								you_view = "유상"
							  else
							  	you_view = "무관"
							end if
                        %>
				        <tr>
				          <td class="first"><%=rs("cost_detail")%></td>
				          <td><a href="overtime_code_mg.asp?overtime_code=<%=rs("overtime_code")%>&u_type=<%="U"%>"><%=rs("work_gubun")%></a></td>
				          <td><%=rs("holi_id")%></td>
				          <td><%=rs("apply_dept")%></td>
				          <td><%=rs("apply_unit")%></td>
				          <td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td>
				          <td><%=rs("meals_yn")%></td>
				          <td><%=rs("work_time1")%>&nbsp;</td>
				          <td><%=rs("sign_yn")%></td>
				          <td><%=you_view%></td>
				          <td><%=rs("use_yn")%></td>
				          <td><%=rs("overtime_memo")%>&nbsp;</td>
			            </tr>
				        <%
							rs.movenext()
						loop
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="1%" valign="top">&nbsp;</td>
				      <td width="24%" valign="top"><form method="post" name="frm" action="overtime_code_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="30%">휴일구분</th>
				              <td class="left">
                                <select name="holi_id" id="holi_id" style="width:150px">
                                  <option value="">선택</option>
                                  <option value='평일' <%If holi_id = "평일" then %>selected<% end if %>>평일</option>
                                  <option value='휴일' <%If holi_id = "휴일" then %>selected<% end if %>>휴일</option>
                                  <option value='스케쥴' <%If holi_id = "스케쥴" then %>selected<% end if %>>스케쥴</option>
                                </select>                        
                              </td>
			                </tr>
				            <tr>
				              <th>수당명</th>
				              <td class="left"><input name="work_gubun" type="text" id="work_gubun" onKeyUp="checklength(this,30)" value="<%=work_gubun%>" style="width:150px"></td>
			                </tr>
				            <tr>
				              <th>적용부서</th>
				              <td class="left"><input name="apply_dept" type="text" id="apply_dept" onKeyUp="checklength(this,30)" value="<%=apply_dept%>" style="width:150px"></td>
			                </tr>
				            <tr>
				              <th>기준</th>
				              <td class="left">
                                <select name="apply_unit" id="apply_unit" style="width:150px">
                                  <option value="">선택</option>
                                  <option value='횟수' <%If apply_unit = "횟수" then %>selected<% end if %>>횟수</option>
                                  <option value='건당' <%If apply_unit = "건당" then %>selected<% end if %>>건당</option>
                                </select>                        
                              </td>
			                </tr>
				            <tr>
				              <th>단가</th>
				              <td class="left"><input name="overtime_amt" type="text" id="overtime_amt" value="<%=formatnumber(overtime_amt,0)%>" onKeyUp="plusComma(this);" style="width:150px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>식대</th>
				              <td class="left">
                              <input type="radio" name="meals_yn" value="포함" <% if meals_yn = "포함" then %>checked<% end if %> style="width:25px" ID="Radio1">포함
				              <input type="radio" name="meals_yn" value="미포함" <% if meals_yn = "미포함" then %>checked<% end if %> style="width:25px" ID="Radio2">미포함
                              </td>
			                </tr>
				            <tr>
				              <th>근무시간</th>
				              <td class="left"><input name="work_time1" type="text" id="work_time1" onKeyUp="checklength(this,20)" value="<%=work_time1%>" style="width:150px"></td>
			                </tr>
				            <tr>
				              <th>사전품의</th>
				              <td class="left">
                              <input type="radio" name="sign_yn" value="Y" <% if sign_yn = "Y" then %>checked<% end if %> style="width:25px" ID="Radio3">Yes
  							  <input type="radio" name="sign_yn" value="N" <% if sign_yn = "N" then %>checked<% end if %> style="width:25px" ID="Radio4">No
							  </td>
			                </tr>
				            <tr>
				              <th>유무상</th>
				              <td class="left">
                              <input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:25px" ID="Radio7">유상
  							  <input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:25px" ID="Radio8">무관
                              </td>
			                </tr>
				            <tr>
				              <th>사용유무</th>
				              <td class="left">
                              <input type="radio" name="use_yn" value="Y" <% if use_yn = "Y" then %>checked<% end if %> style="width:25px" ID="Radio5">사용
  							  <input type="radio" name="use_yn" value="N" <% if use_yn = "N" then %>checked<% end if %> style="width:25px" ID="Radio6">미사용
                              </td>
			                </tr>
				            <tr>
				              <th>비고</th>
				              <td class="left"><input name="overtime_memo" type="text" id="overtime_memo" onKeyUp="checklength(this,50)" value="<%=overtime_memo%>" style="width:150px"></td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="overtime_code" value="<%=overtime_code%>" ID="Hidden1">
				        <input type="hidden" name="work_time2" value="<%=work_time2%>" ID="Hidden1">
				        <div align=center>
                        	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        	<span class="btnType01"><input type="button" value="취소" onclick="javascript:frmcancel();" ID="Button1" NAME="Button1"></span>
                        </div>
			          </form></td>
			        </tr>
				    <tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

