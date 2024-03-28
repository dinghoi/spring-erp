<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

emp_no = request("emp_no")
emp_name = request("emp_name")

in_empno =""
in_name = ""
If Request.Form("in_name")  <> "" Then 
  in_empno = Request.Form("in_empno") 
  in_name = Request.Form("in_name") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

    in_pay_sum = 0 
'	give_pay_sum = 0
	
    sql="select * from emp_sawo_in where in_empno = '" + emp_no + "'"
	Rs_sum.Open Sql, Dbconn, 1
	
	do until rs_sum.eof
	   in_pay_sum = in_pay_sum + rs_sum("in_pay")
'	   give_pay_sum = give_pay_sum + rs_sum("sawo_give_pay")
	   
	   rs_sum.movenext()
	loop
    rs_sum.close()

sql = "select * from emp_sawo_in where in_empno = '" + emp_no + "' ORDER BY in_empno,in_date,in_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "경조금 납부 내역〔개인별〕"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('성명을 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_in_view.asp?emp_no=<%=emp_no%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:100px; text-align:left">
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=emp_name%>" style="width:150px; text-align:left">
								</label>
                            <strong>총 납입액 : </strong>
                                <label>
                               	<input name="in_pay_sum" type="text" id="in_pay_sum" value="<%=formatnumber(clng(in_pay_sum),0)%>" style="width:150px; text-align:right">
								</label>    
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="14%" >
							<col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">납부일자</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">납부금액</th>
                                <th scope="col">비고</th>
 							</tr>
						</thead>
						<tbody>
						<%
						    v_cnt = 0
							do until rs.eof or rs.bof
							    v_cnt = v_cnt + 1
						%>
							<tr>
								<td class="first"><%=rs("in_date")%>&nbsp;</td>
                                <td><%=rs("in_company")%>&nbsp;</td>
                                <td><%=rs("in_org_name")%>&nbsp;</td>
                                <td><%=rs("in_pay")%>&nbsp;</td>
                                <td><%=rs("in_comment")%>&nbsp;</td>                                
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()

						    if v_cnt = 0 then 
						%>
							<tr>
								<td class="first" colspan="5">내역이 없습니다</td>
							</tr>
                        <%     else %>
                           	<tr>
								<td class="first" colspan="5"><%=v_cnt%>건이 조회되었습니다</td>
							</tr>
						<%  end if %>
						</tbody>
					</table>
				</div>
			</div>				
	   </div>     
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>       				
	</form>
	</body>
</html>

