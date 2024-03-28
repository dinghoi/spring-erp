<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
fm_id = request("fm_id")
fm_type = request("fm_type")

fm_sawo_pay = 0
fm_company_pay = 0
fm_holiday1 = ""
fm_holiday2 = ""
fm_wreath_yn = ""
fm_flowers_yn = ""
fm_comment = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 경조금 지급규정 등록 "
if u_type = "U" then

	Sql="select * from emp_family_event where fm_id = '"&fm_id&"' and fm_type = '"&fm_type&"'"
	Set rs=DbConn.Execute(Sql)

	fm_sawo_pay = rs("fm_sawo_pay")
    fm_company_pay = rs("fm_company_pay")
    fm_holiday1 = rs("fm_holiday1")
    fm_holiday2 = rs("fm_holiday2")
    fm_wreath_yn = rs("fm_wreath_yn")
    fm_flowers_yn = rs("fm_flowers_yn")
    fm_comment = rs("fm_comment")
	
	rs.close()

	title_line = " 경조금 지급규정 변경 "
	
end if

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
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				sawo_pay = parseInt(document.frm.fm_sawo_pay.value.replace(/,/g,""));		
				sawo_pay = String(sawo_pay);
				num_len = sawo_pay.length;
				sil_len = num_len;
				sawo_pay = String(sawo_pay);
				if (sawo_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) sawo_pay = sawo_pay.substr(0,num_len -3) + "," + sawo_pay.substr(num_len -3,3);
				if (sil_len > 6) sawo_pay = sawo_pay.substr(0,num_len -6) + "," + sawo_pay.substr(num_len -6,3) + "," + sawo_pay.substr(num_len -2,3);
				document.frm.fm_sawo_pay.value = sawo_pay; 
				
				company_pay = parseInt(document.frm.fm_company_pay.value.replace(/,/g,""));		
				company_pay = String(company_pay);
				num_len = company_pay.length;
				sil_len = num_len;
				company_pay = String(company_pay);
				if (company_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) company_pay = company_pay.substr(0,num_len -3) + "," + company_pay.substr(num_len -3,3);
				if (sil_len > 6) company_pay = company_pay.substr(0,num_len -6) + "," + company_pay.substr(num_len -6,3) + "," + company_pay.substr(num_len -2,3);
				document.frm.fm_company_pay.value = company_pay; 
			}									
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_fm_event_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">경조구분</th>
                      <td colspan="2" class="left" bgcolor="#FFFFE6">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '11' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="fm_id" id="fm_id" style="width:130px">
                         <option value="" <% if fm_id = "" then %>selected<% end if %>>선택</option>
                	<% 
					  do until rs_etc.eof 
		            %>
                	     <option value='<%=rs_etc("emp_etc_name")%>' <%If fm_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                    <%
						rs_etc.movenext()  
					  loop 
					  rs_etc.Close()
					%>
            		  </select>    
                      </td>
                      <th style="background:#FFFFE6">경조유형</th>
                      <td colspan="2" class="left" bgcolor="#FFFFE6">
					<%
					  Sql="select * from emp_etc_code where emp_etc_type = '12' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="fm_type" id="fm_type" style="width:130px">
                         <option value="" <% if fm_type = "" then %>selected<% end if %>>선택</option>
                	<% 
					  do until rs_etc.eof 
		            %>
                	     <option value='<%=rs_etc("emp_etc_name")%>' <%If fm_type = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                    <%
						rs_etc.movenext()  
					  loop 
					  rs_etc.Close()
					%>
            		  </select>    
                      </td>
                    </tr>
                    <tr>
                      <th>경조회<br>경조금</th>
                      <td colspan="5" class="left">
                      <input name="fm_sawo_pay" type="text" id="fm_sawo_pay" style="width:90px;text-align:right" value="<%=formatnumber(clng(fm_sawo_pay),0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>회사<br>축의금</th>
                      <td colspan="5" class="left">
                      <input name="fm_company_pay" type="text" id="fm_company_pay" style="width:90px;text-align:right" value="<%=formatnumber(clng(fm_company_pay),0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>휴가일수<br>유급</th>
                      <td colspan="2" class="left">
					  <input name="fm_holiday1" type="text" id="fm_holiday1" size="3" maxlength="1" value="<%=fm_holiday1%>"></td>
                      <th>휴가일수<br>무급</th>
                      <td colspan="2" class="left">
                      <input name="fm_holiday2" type="text" id="fm_holiday2" size="3" maxlength="1" value="<%=fm_holiday2%>"></td>
                    </tr>
                    <tr>
                      <th>화환<br>조화</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="fm_wreath_yn" value="Y" <% if fm_wreath_yn = "Y" then %>checked<% end if %>>지급 
              		  <input name="fm_wreath_yn" type="radio" value="N" <% if fm_wreath_yn = "N" then %>checked<% end if %>>안함
					  </td>
                      <th>꽃다발</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="fm_flowers_yn" value="Y" <% if fm_flowers_yn = "Y" then %>checked<% end if %>>지급 
              		  <input name="fm_flowers_yn" type="radio" value="N" <% if fm_flowers_yn = "N" then %>checked<% end if %>>안함
					  </td>
                    </tr>
                    <tr>
                      <th>비고</th>  
                      <td colspan="5" class="left">
					  <input name="fm_comment" type="text" id="fm_comment" style="width:200px" value="<%=fm_comment%>">
					  </td>
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

