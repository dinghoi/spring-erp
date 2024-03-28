<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
'ask_empno = request("sawo_empno")
in_seq = request("in_seq")
emp_name = request("emp_name")

in_name = request.cookies("nkpmg_user")("coo_user_name")
ask_empno = request.cookies("nkpmg_user")("coo_user_id")


ask_company = ""
ask_org = ""
ask_org_name = ""
ask_id = ""
ask_type = ""
ask_sawo_place = ""
ask_sawo_comm = ""
ask_att_file = ""
ask_process = "0"
att_file = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 경조금 신청 "
if u_type = "U" then

	Sql="select * from emp_sawo_ask where ask_empno = '"&ask_empno&"' and ask_seq = '"&in_seq&"' and ask_date = '"&ask_date&"'"
	Set rs=DbConn.Execute(Sql)

	ask_empno = rs("ask_empno")
    ask_seq = rs("ask_seq")
    ask_date = rs("ask_date")
    ask_emp_name = rs("ask_emp_name")
    ask_company = rs("ask_company")
    ask_org = rs("ask_org")
    ask_org_name = rs("ask_org_name")
	ask_id = rs("ask_id")
    ask_type = rs("giveask_type_type")
    ask_sawo_place = rs("ask_sawo_place")
    ask_sawo_comm = rs("ask_sawo_comm")
	ask_att_file = rs("ask_att_file")
	
	rs.close()

	title_line = " 경조금 신청 변경 "
	
end if

    sql="select max(ask_seq) as max_seq from emp_sawo_ask where ask_empno = '"&ask_empno&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "001"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,3)
	end if
    rs_max.close()
	
	if u_type = "U" then
	   code_last = ask_seq
	end if
	
ask_seq = code_last

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
												$( "#datepicker" ).datepicker("setDate", "<%=ask_date%>" );
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
				if(document.frm.ask_id.value =="") {
					alert('경조구분을 선택하세요');
					frm.ask_id.focus();
					return false;}
				if(document.frm.ask_type =="") {
					alert('경조휴형을 선택하세요');
					frm.ask_type.focus();
					return false;}
				if(document.frm.ask_date.value =="") {
					alert('경조발생일를 입력하세요');
					frm.ask_date.focus();
					return false;}
				if(document.frm.ask_sawo_place.value =="") {
					alert('경조발생장소를 입력하세요');
					frm.ask_sawo_place.focus();
					return false;}
				if(document.frm.ask_sawo_comm.value =="") {
					alert('경조 기타사항을 입력하세요');
					frm.ask_sawo_comm.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function file_browse()	{ 
           		document.frm.att_file.click(); 
           		document.frm.text1.value=document.frm.att_file.value;  
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_ask_save.asp" method="post" name="frm" enctype="multipart/form-data">
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
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="ask_empno" type="text" id="ask_empno" size="14" value="<%=ask_empno%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="ask_emp_name" type="text" id="ask_emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    <%
                         if ask_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&ask_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
							  emp_org_code = Rs_emp("emp_org_code")
							  emp_org_name = Rs_emp("emp_org_name")
							  emp_company = Rs_emp("emp_company")
							  emp_bonbu = Rs_emp("emp_bonbu")
							  emp_saupbu = Rs_emp("emp_saupbu")
							  emp_team = Rs_emp("emp_team")
							  emp_reside_place = Rs_emp("emp_reside_place")
		                   end if
	                       Rs_emp.Close()
	                	  end if	
				    %>	
                      <th style="background:#FFFFE6">직급/직책</th>                      
                      <td class="left" bgcolor="#FFFFE6"><%=emp_grade%>&nbsp;-&nbsp;<%=emp_position%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>소속</th>                      
                      <td class="left"><%=emp_org_code%>&nbsp;-&nbsp;<%=emp_org_name%>&nbsp;</td>
                      <th>조직</th>                      
                      <td colspan="3" class="left"><%=emp_company%>&nbsp;-&nbsp;<%=emp_bonbu%>&nbsp;-&nbsp;<%=emp_saupbu%>&nbsp;-&nbsp;<%=emp_team%>&nbsp;</td>
                    </tr>
                 	<tr>
                      <th>경조구분</th>
                      <td class="left">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '11' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="ask_id" id="ask_id" style="width:130px">
                         <option value="" <% if ask_id = "" then %>selected<% end if %>>선택</option>
                	<% 
					  do until rs_etc.eof 
		            %>
                	     <option value='<%=rs_etc("emp_etc_name")%>' <%If ask_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                    <%
						rs_etc.movenext()  
					  loop 
					  rs_etc.Close()
					%>
            		  </select>  
                      </td>
                      <th>경조유형</th>
                      <td colspan="3" class="left">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '12' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="ask_type" id="ask_type" style="width:130px">
                         <option value="" <% if ask_type = "" then %>selected<% end if %>>선택</option>
                	<% 
					  do until rs_etc.eof 
		            %>
                	     <option value='<%=rs_etc("emp_etc_name")%>' <%If ask_type = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                    <%
						rs_etc.movenext()  
					  loop 
					  rs_etc.Close()
					%>
            		  </select>    
                      </td>
                    </tr>
                    <tr>
                      <th>경조일시</th>
                      <td colspan="5" class="left">
					  <input name="ask_date" type="text" value="<%=ask_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                    </tr>
                    <tr>
                      <th>경조장소</th>
                      <td colspan="5" class="left">
					  <input name="ask_sawo_place" type="text" id="ask_sawo_place" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=ask_sawo_place%>"></td>
                    </tr>
                    <tr>
                      <th>기타</th>
                      <td colspan="5" class="left">
					  <input name="ask_sawo_comm" type="text" id="ask_sawo_comm" style="width:300px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=ask_sawo_comm%>">
                      &nbsp;&nbsp;(연락처, 발인일, 발인장소 등)
                      </td>
                    </tr>
                    <tr>
                      <th>No.</th>  
					  <td colspan="5" class="left"><%=ask_seq%><input name="ask_seq" type="hidden" value="<%=ask_seq%>"></td>
			    	</tr>
                    <tr>
                      <th scope="row">첨부파일</th>  
					  <td colspan="5" class="left">
					  <input type="file" name= "att_file"  size="70" accept="image/gif"><br> * 첨부파일은 1개만 가능하며 최대용량은 2MB
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
                <input type="hidden" name="ask_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="ask_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="ask_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="ask_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="ask_org" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="ask_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                <input type="hidden" name="v_att_file" value="<%=att_file%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

