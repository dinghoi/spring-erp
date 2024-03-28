<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
stay_code = request("stay_code")
stay_name = request("stay_name")

code_last = ""

stay_org_code = ""
stay_name = ""
stay_sido = ""
stay_gugun = ""
stay_dong = ""
stay_addr = ""
stay_tel_ddd = ""
stay_tel_no1 = ""
stay_tel_no2 = ""
stay_fax_ddd = ""
stay_fax_no1 = ""
stay_fax_no2 = ""
stay_reg_date = ""

curr_date = mid(cstr(now()),1,10)
'view_condi = "(주)케이원정보통신"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "◈ 실근무지 등록 ◈"
if u_type = "U" then

	Sql="select * from emp_stay where stay_code = '"&stay_code&"'"
	Set rs=DbConn.Execute(Sql)

	stay_name = rs("stay_name")
    stay_org_code = rs("stay_org_code")
	stay_org_name = rs("stay_org_name")
	stay_reside_company = rs("stay_reside_company")
    stay_sido = rs("stay_sido")
    stay_gugun = rs("stay_gugun")
    stay_dong = rs("stay_dong")
    stay_addr = rs("stay_addr")
    stay_tel_ddd = rs("stay_tel_ddd")
    stay_tel_no1 = rs("stay_tel_no1")
    stay_tel_no2 = rs("stay_tel_no2")
    stay_fax_ddd = rs("stay_fax_ddd")
    stay_fax_no1 = rs("stay_fax_no1")
    stay_fax_no2 = rs("stay_fax_no2")
    stay_reg_date = rs("stay_reg_date")
	
	rs.close()

	title_line = "◈ 실근무지 변경 ◈"
	
end if

    sql="select max(stay_code) as max_seq from emp_stay"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "0001"
	  else
		max_seq = "000" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,4)
	end if
    rs_max.close()
	
	if u_type = "U" then
	   code_last = stay_code
	end if
	
stay_code = code_last
'response.write(org_code)


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
												$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
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
				if(document.frm.stay_name.value =="") {
					alert('근무지명을 입력하세요');
					frm.stay_name.focus();
					return false;}
				if(document.frm.stay_sido =="") {
					alert('주소를 선택하세요');
					frm.stay_sido.focus();
					return false;}
				if(document.frm.stay_addr.value =="") {
					alert('번지를 입력하세요');
					frm.stay_addr.focus();
					return false;}
				if(document.frm.stay_tel_no1.value =="") {
					alert('전화번호를를 입력하세요');
					frm.stay_tel_no1.focus();
					return false;}
				
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
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_stay_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
                        <col width="8" >
						<col width="22%" >
						<col width="8%" >
						<col width="22%" >
						<col width="8%" >
						<col width="22" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th class="left" style="background:#FFFFE6">코&nbsp;&nbsp;&nbsp;&nbsp;드</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6">
					  <input name="stay_code" type="text" id="stay_code" size="4" value="<%=stay_code%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th class="left">실근무지명</th>
                      <td class="left">
					  <input name="stay_name" type="text" id="stay_name" size="16" value="<%=stay_name%>"></td>
                      <th>회사</th>
                      <td colspan="3" class="left">
                    <%
					  Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                  rs_org.Open Sql, Dbconn, 1
					%>
					  <select name="view_condi" id="view_condi" type="text" style="width:150px">
                         <option value="" <% if view_condi = "" then %>selected<% end if %>>선택</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                				<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            		  </select>                             
                      </td>                                            
                    </tr>
                    <tr>
                      <th class="left">상주처명</th>
                      <td colspan="2" class="left">
					  <input name="stay_org_name" type="text" id="stay_org_name" size="16" readonly="true" value="<%=stay_org_name%>">
                      <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="stay"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
                      </td>
                      <th>상주처 회사</th>
                      <td colspan="2" class="left">
                      <input name="stay_reside_company" type="text" id="stay_reside_company" size="16" readonly="true" value="<%=stay_reside_company%>">
                      <input name="stay_org_code" type="hidden" id="stay_org_code" value="<%=stay_org_code%>">
                      </td>
                    </tr>
                 	<tr>  
                      <th class="left">주&nbsp;&nbsp;&nbsp;&nbsp;소</th>
					  <td colspan="5" class="left">
                      <input name="stay_sido" type="text" id="stay_sido" style="width:100px" readonly="true" value="<%=stay_sido%>">
              		  <input name="stay_gugun" type="text" id="stay_gugun" style="width:150px" readonly="true" value="<%=stay_gugun%>">
              		  <input name="stay_dong" type="text" id="stay_dong" style="width:150px" readonly="true" value="<%=stay_dong%>">
              		  <input name="stay_zip" type="hidden" id="stay_zip" value="">
                      <a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="stay"%>','stay_zip_select','scrollbars=yes,width=600,height=400')">주소조회</a>
                      </td>
                    </tr>
                 	<tr>  
                      <th class="left">번&nbsp;&nbsp;&nbsp;&nbsp;지</th>
					  <td colspan="5" class="left">
              		  <input name="stay_addr" type="text" id="stay_addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=stay_addr%>">
                      </td>
                    </tr>                    
                 	<tr>     
					  <th class="left">전화번호</th>
                      <td colspan="5" class="left">
                      <input name="stay_tel_ddd" type="text" id="stay_tel_ddd" size="3" maxlength="3" value="<%=stay_tel_ddd%>" >
					  -
                      <input name="stay_tel_no1" type="text" id="stay_tel_no1" size="4" maxlength="4" value="<%=stay_tel_no1%>" >
                      -
                      <input name="stay_tel_no2" type="text" id="stay_tel_no2" size="4" maxlength="4" value="<%=stay_tel_no2%>" ></td>                    </tr>
                 	<tr>     
					  <th class="left">팩스번호</th>
                      <td colspan="5" class="left">
                      <input name="stay_fax_ddd" type="text" id="stay_fax_ddd" size="3" maxlength="3" value="<%=stay_fax_ddd%>" >
					  -
                      <input name="stay_fax_no1" type="text" id="stay_fax_no1" size="4" maxlength="4" value="<%=stay_fax_no1%>" >
                      -
                      <input name="stay_fax_no2" type="text" id="stay_fax_no2" size="4" maxlength="4" value="<%=stay_fax_no2%>" ></td>  
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

