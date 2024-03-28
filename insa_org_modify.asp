<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
org_code = request("org_code")

code_last = ""

org_level = ""
org_name = ""
org_date = ""
org_end_date = ""
org_empno = ""
org_empname = ""
org_company = ""
org_bonbu = ""
org_saupbu = ""
org_team = ""
org_reside_place = ""
org_reside_company = ""
org_cost_group = ""
owner_org = ""
owner_orgname = ""
owner_empno = ""
owner_empname = ""
org_table_org = 0
tel_ddd = ""
tel_no1 = ""
tel_no2 = ""
org_sido = ""
org_gugun = ""
org_dong = ""
org_addr = ""
org_end_date = ""
org_reg_date = ""
org_reg_user = ""
org_mod_date = ""
org_mod_user = ""

' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_tra = Server.CreateObject("ADODB.Recordset")
Set Rs_owner = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'response.write u_type&"<br>"
if u_type = "U" then

	Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	'response.write Sql&"<br>"
	Set rs=DbConn.Execute(Sql)

    org_level = rs("org_level")
    org_name = rs("org_name")
    org_date = rs("org_date")
	org_end_date = rs("org_end_date")
    org_empno = rs("org_empno")
    org_empname = rs("org_emp_name")
    org_company = rs("org_company")
    org_bonbu = rs("org_bonbu")
    org_saupbu = rs("org_saupbu")
    org_team = rs("org_team")
	org_reside_place = rs("org_reside_place")
	org_reside_company = rs("org_reside_company")
	org_cost_group = rs("org_cost_group")
	org_cost_center = rs("org_cost_center")
    owner_org = rs("org_owner_org")
    owner_empno = rs("org_owner_empno")
    owner_empname = rs("org_owner_empname")
	if rs("org_table_org") = "" or isnull(rs("org_table_org")) then
	        org_table_org = 0
	   else
			org_table_org = rs("org_table_org")
	end if
    tel_ddd = rs("org_tel_ddd")
    tel_no1 = rs("org_tel_no1")
    tel_no2 = rs("org_tel_no2")
	org_sido = rs("org_sido")
    org_gugun = rs("org_gugun")
    org_dong = rs("org_dong")
    org_addr = rs("org_addr")
    org_reg_date = rs("org_reg_date")
	org_reg_user = rs("org_reg_user")
    org_mod_date = rs("org_mod_date")
    org_mod_user = rs("org_mod_user")
	rs.close()

	Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
	Set rs_owner=DbConn.Execute(Sql)

    owner_orgname = rs_owner("org_name")
	rs_owner.close()

	title_line = " 조직의 상주처(회사) 및 기타정보 수정 "
end if


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=org_date%>" );
			});
			$(function() {  $( "#datepicker1" ).datepicker();
							$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker1" ).datepicker("setDate", "<%=org_end_date%>" );
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
				if(document.frm.org_code.value =="") {
					alert('조직코드를 입력하세요');
					frm.org_code.focus();
					return false;}
				if(document.frm.org_name.value =="") {
					alert('조직명을 입력하세요');
					frm.org_name.focus();
					return false;}
				if(document.frm.org_date.value =="") {
					alert('조직생성일을 입력하세요');
					frm.org_date.focus();
					return false;}
				if(document.frm.org_empno.value =="") {
					alert('조직장사번을 입력하세요');
					frm.org_empno.focus();
					return false;}
				if(document.frm.org_empname.value =="") {
					alert('조직장성명을 입력하세요');
					frm.org_empname.focus();
					return false;}
//				if(document.frm.org_cost_group.value =="") {
//					alert('비용센타그룹을 선택하세요');
//					frm.org_cost_group.focus();
//					return false;}
				if(document.frm.org_cost_center.value ="상주직접비") {
					if(document.frm.org_reside_place.value =="") {
						alert('상주처를 입력하세요');
						frm.org_reside_place.focus();
						return false;}}
				if(document.frm.org_cost_center.value =="상주직접비") {
					if(document.frm.org_reside_company.value =="") {
						alert('상주처회사를 입력하세요');
						frm.org_reside_company.focus();
						return false;}}
				if(document.frm.org_level.value !="회사") {
					if(document.frm.owner_org.value =="") {
						alert('상위조직을 입력하세요');
						frm.owner_org.focus();
						return false;}}
				if(document.frm.org_level.value =="상주처") {
					if(document.frm.org_reside_company.value =="") {
						alert('상주처 회사를 선택하세요');
						frm.org_reside_company.focus();
						return false;}}
				if(document.frm.org_level.value =="상주처") {
					if(document.frm.org_cost_center.value !="상주직접비") {
						alert('비용구분에 상주직접비를 선택하세요');
						frm.org_cost_center.focus();
						return false;}}
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function num_chk(txtObj){
				org_to = parseInt(document.frm.org_table_org.value.replace(/,/g,""));

				org_to = String(org_to);
				num_len = org_to.length;
				sil_len = num_len;
				org_to = String(org_to);
				if (org_to.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) org_to = org_to.substr(0,num_len -3) + "," + org_to.substr(num_len -3,3);
				if (sil_len > 6) org_to = org_to.substr(0,num_len -6) + "," + org_to.substr(num_len -6,3) + "," + org_to.substr(num_len -2,3);
				document.frm.org_table_org.value = org_to;

			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">

			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_org_reg_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">조직코드</th>
                                <td class="left"><%=org_code%><input name="org_code" type="hidden" value="<%=org_code%>"></td>
                                <th>조직명</th>
                                <td class="left"><input name="org_name" type="text" readonly="true" value="<%=org_name%>"></td>
                                <th>조직&nbsp;구분</th>
                                <td class="left"><input name="org_level" type="text" readonly="true" value="<%=org_level%>"></td>
                                <th>조직생성일</th>
                                <td class="left">
                                <input name="org_date" type="text" size="10" readonly="true" value="<%=org_date%>" >
              					</td>
                             </tr>
                             <tr>
								<th class="first">조직장사번</th>
                                <td class="left"><input name="org_empno" type="text" id="org_empno" size="7" readonly="true" value="<%=org_empno%>">
                                </td>
                                <th>조직장성명</th>
                                <td class="left">
                                <input name="org_empname" type="text" id="org_empname" size="10" readonly="true" value="<%=org_empname%>">
                                </td>
                                <th>소속</th>
                                <td colspan="3" class="left">
                                <input name="org_company" type="text" id="org_company" style="width:100px" readonly="true" value="<%=org_company%>">
              					<input name="org_bonbu" type="text" id="org_bonbu" style="width:120px" readonly="true" value="<%=org_bonbu%>">
              					<input name="org_saupbu" type="text" id="org_saupbu" style="width:120px" readonly="true" value="<%=org_saupbu%>">
              					<input name="org_team" type="text" id="org_team" style="width:120px" readonly="true" value="<%=org_team%>">
                                </td>
                             </tr>
							<tr>
								<th class="first">상위조직코드</th>
                                <td class="left"><input name="owner_org" type="text" id="owner_org" size="4" readonly="true" value="<%=owner_org%>">
                                </td>
                                <th>상위조직명</th>
                                <td class="left">
                                <input name="owner_orgname" type="text" id="owner_orgname" size="20" readonly="true" value="<%=owner_orgname%>"></td>
                                <th>상위조직장</th>
                                <td class="left">
                                <input name="owner_empno" type="text" id="owner_empno" size="7" readonly="true" value="<%=owner_empno%>"></td>
                                <th>상위조직장명</th>
                                <td class="left">
                                <input name="owner_empname" type="text" id="owner_empname" size="20" readonly="true" value="<%=owner_empname%>"></td>
                             </tr>
                             <tr>
								<th class="first" style="background:#FFC">대표전화</th>
                                <td class="left"><input name="tel_ddd" type="text" id="tel_ddd" size="3" maxlength="3" value="<%=tel_ddd%>" >
								  -
                                    <input name="tel_no1" type="text" id="tel_no1" size="4" maxlength="4" value="<%=tel_no1%>" >
                                    -
                                <input name="tel_no2" type="text" id="tel_no2" size="4" maxlength="4" value="<%=tel_no2%>" ></td>
                                <th>조직폐쇄일</th>
                                <td class="left">
                                <input name="org_end_date" type="text" size="10" readonly="true" id="datepicker1" style="width:70px;" value="<%=org_end_date%>" >
              					</td>
                                <th style="background:#FFC">상주처</th>
                                <td class="left"><input name="org_reside_place" type="text" id="org_reside_place" style="width:150px" value="<%=org_reside_place%>">
                                </td>
                                <th class="first" style="background:#FFC">상주처 회사</th>
								<td class="left"><input name="org_reside_company" type="text" id="org_reside_company" style="width:120px" readonly="true" value="<%=org_reside_company%>">
								<a href="#" class="btnType03" onClick="pop_Window('insa_trade_search.asp?gubun=<%="1"%>','tradesearch','scrollbars=yes,width=600,height=400')">찾기</a>
            					</td>

                             </tr>
                             <tr>
								<th class="first" style="background:#FFC">주소</th>
								<td colspan="5" class="left">
                                <input name="org_sido" type="text" id="org_sido" style="width:100px" readonly="true" value="<%=org_sido%>">
              					<input name="org_gugun" type="text" id="org_gugun" style="width:150px" readonly="true" value="<%=org_gugun%>">
              					<input name="org_dong" type="text" id="org_dong" style="width:150px" readonly="true" value="<%=org_dong%>">
              					<input name="org_addr" type="text" id="org_addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=org_addr%>">
              					<input name="org_zip" type="hidden" id="org_zip" value="">
                                <a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="org"%>','org_zip_select','scrollbars=yes,width=600,height=400')">주소조회</a>
                                </td>
                                <th>비용센타그룹</th>
                                <td class="left"><input name="org_cost_group" type="text" id="org_cost_group" style="width:150px" readonly="true" value="<%=org_cost_group%>">
            					</td>
                              </tr>
                              <tr>
								<th class="first" style="background:#FFC">적정인원(T.O)</th>
								<td class="left">
                                <input name="org_table_org" type="text" id="org_table_org" style="width:90px;text-align:right" value="<%=formatnumber(org_table_org,0)%>" onKeyUp="num_chk(this);">
            					</td>
                                <th>입력일자</th>
                                <td class="left">
                                <input name="org_reg_date" type="text" id="org_reg_date" style="width:150px" readonly="true" value="<%=org_reg_date%>">
                                </td>
                                <th>수정일자</th>
                                <td class="left">
                                <input name="org_mod_date" type="text" id="org_mod_date" style="width:150px" readonly="true" value="<%=org_mod_date%>">
                                </td>
                                <th style="background:#FFC">비용구분</th>
                                <td class="left">
                                    <%
                                    Sql="select * from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
                                    Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                    <select name="org_cost_center" id="org_cost_center" style="width:90px">
                                        <option value="" <% if org_cost_center = "" then %>selected<% end if %>>선택</option>
                                        <%
                                        do until rs_etc.eof
                                        %>
                                        <option value='<%=rs_etc("emp_etc_name")%>' <%If org_cost_center = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                                        <%
                                            rs_etc.movenext()
                                        loop
                                        rs_etc.Close()
                                        %>
                			        </select>
                                </td>
                              </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <% if u_type = "U" then %>
					      <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
					<%   else  %>
                         <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"></span>
                    <% end if %>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="mg_level" value="<%=org_level%>" ID="Hidden1">
				</form>
		</div>
	</div>
	</body>
</html>

