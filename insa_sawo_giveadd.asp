<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
sawo_empno = request("sawo_empno")
in_seq = request("in_seq")
emp_name = request("emp_name")
ask_seq = request("ask_seq")
ask_date = request("ask_date")
give_ask_process = request("give_ask_process")

'response.write(give_ask_process)
'response.write(ask_date)

give_company = ""
give_org = ""
give_org_name = ""
give_id = ""
give_type = ""
give_pay = 0
give_sawo_date = ""
give_sawo_place = ""
give_sawo_comm = ""
give_comment = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_ask = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if give_ask_process = "1" then 
      title_line = " 경조금 지급 등록 "
   else
      title_line = " 경조회 경조금 지급 등록 "
end if

if u_type = "U" then

	Sql="select * from emp_sawo_give where give_empno = '"&sawo_empno&"' and give_seq = '"&in_seq&"' and give_date = '"&give_date&"' and give_ask_process = '"&give_ask_process&"'"
	Set rs=DbConn.Execute(Sql)

	give_empno = rs("give_empno")
    give_seq = rs("give_seq")
    give_date = rs("give_date")
	give_ask_process = rs("give_ask_process")
    give_emp_name = rs("give_emp_name")
    give_company = rs("give_company")
    give_org = rs("give_org")
    give_org_name = rs("give_org_name")
    give_pay = rs("give_pay")
	give_comment = rs("give_comment")
	give_id = rs("give_id")
    give_type = rs("give_type")
    give_sawo_date = rs("give_sawo_date")
    give_sawo_place = rs("give_sawo_place")
    give_sawo_comm = rs("give_sawo_comm")
	
	rs.close()

	title_line = " 경조금 지급 변경 "
	
end if

    sql="select max(give_seq) as max_seq from emp_sawo_give where give_empno = '"&sawo_empno&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "001"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,3)
	end if
    rs_max.close()
	
	if u_type = "U" then
	   code_last = give_seq
	end if
	
give_seq = code_last

sql = "select * from emp_sawo_ask  where ask_empno = '"&sawo_empno&"' and ask_seq = '"&ask_seq&"' and ask_date = '"&ask_date&"'"
Rs_ask.Open Sql, Dbconn, 1

'response.write(rs_ask("ask_id"))
    give_company = rs_ask("ask_company")
    give_org = rs_ask("ask_org")
    give_org_name = rs_ask("ask_org_name")
	give_id = rs_ask("ask_id")
    give_type = rs_ask("ask_type")
    give_sawo_date = rs_ask("ask_date")
    give_sawo_place = rs_ask("ask_sawo_place")
    give_sawo_comm = rs_ask("ask_sawo_comm")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=give_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=give_sawo_date%>" );
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
				if(document.frm.give_date.value =="") {
					alert('경조금지급일을 입력하세요');
					frm.give_date.focus();
					return false;}
				if(document.frm.give_pay.value =="") {
					alert('경조금액을 입력하세요');
					frm.give_pay.focus();
					return false;}
								
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function sawo_cal(txtObj){
				sawo_amt = parseInt(document.frm.give_pay.value.replace(/,/g,""));		
				sawo_amt = String(sawo_amt);
				num_len = sawo_amt.length;
				sil_len = num_len;
				sawo_amt = String(sawo_amt);
				if (sawo_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) sawo_amt = sawo_amt.substr(0,num_len -3) + "," + sawo_amt.substr(num_len -3,3);
				if (sil_len > 6) sawo_amt = sawo_amt.substr(0,num_len -6) + "," + sawo_amt.substr(num_len -6,3) + "," + sawo_amt.substr(num_len -2,3);

				document.frm.give_pay.value = sawo_amt; 

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
			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_giveadd_save.asp" method="post" name="frm">
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
                      <td class="left" bgcolor="#FFFFE6"><%=sawo_empno%></td>
					  <input name="give_empno" type="hidden" id="give_empno" size="14" value="<%=sawo_empno%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6"><%=emp_name%></td>
					  <input name="give_emp_name" type="hidden" id="give_emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    <%
                         if sawo_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&sawo_empno&"'"
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
                      <td class="left"><%=rs_ask("ask_id")%>&nbsp;</td>
                      <th>경조유형</th>
					  <td colspan="3" class="left"><%=rs_ask("ask_type")%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>지급일</th>
                      <td colspan="5" class="left">
					  <input name="give_date" type="text" value="<%=give_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                    </tr>
                    <tr>
                      <th>경조금액</th>
                      <td colspan="5" class="left">
					  <input name="give_pay" type="text" id="give_pay" style="width:80px;text-align:right" value="<%=formatnumber(clng(give_pay),0)%>" onKeyUp="sawo_cal(this);"></td>
                    </tr>
                    <tr>
                      <th>경조일시</th>
                      <td colspan="5" class="left"><%=rs_ask("ask_date")%>&nbsp;</td>
                    <tr>
                      <th>경조장소</th>
                      <td colspan="5" class="left"><%=rs_ask("ask_sawo_place")%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>기타</th>
                      <td colspan="5" class="left"><%=rs_ask("ask_sawo_comm")%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>경조<br>Comment.</th>
                      <td colspan="5" class="left">
					  <input name="give_comment" type="text" id="give_comment" style="width:500px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=give_comment%>">
                      </td>
                    </tr>
                    <tr>
                      <th>No.</th>  
					  <td colspan="5" class="left"><%=give_seq%><input name="give_seq" type="hidden" value="<%=give_seq%>"></td>
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
                <input type="hidden" name="give_company" value="<%=give_company%>" ID="Hidden1">
                <input type="hidden" name="give_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="give_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="give_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="give_org" value="<%=give_org%>" ID="Hidden1">
                <input type="hidden" name="give_org_name" value="<%=give_org_name%>" ID="Hidden1">
                <input type="hidden" name="give_id" value="<%=give_id%>" ID="Hidden1">
                <input type="hidden" name="give_type" value="<%=give_type%>" ID="Hidden1">
                <input type="hidden" name="give_sawo_date" value="<%=give_sawo_date%>" ID="Hidden1">
                <input type="hidden" name="give_sawo_place" value="<%=give_sawo_place%>" ID="Hidden1">
                <input type="hidden" name="give_sawo_comm" value="<%=give_sawo_comm%>" ID="Hidden1">
                <input type="hidden" name="ask_seq" value="<%=ask_seq%>" ID="Hidden1">
                <input type="hidden" name="give_ask_process" value="<%=give_ask_process%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

