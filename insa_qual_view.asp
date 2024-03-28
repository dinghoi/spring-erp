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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

'if emp_name = "" then
'	first_view = "N"
'	sql = "select * from emp_career where career_empno = '" + emp_no + "'"
' else
'	first_view = "Y"
'	sql = "select * from emp_career where  career_empno like '%" + emp_no + "%' ORDER BY career_empno ASC"
'end if

Sql = "SELECT count(*) FROM emp_qual where qual_empno = '"&emp_no&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

if tottal_record = 0 then
   first_view = "Y"
end if

sql = "select * from emp_qual where qual_empno = '" + emp_no + "' ORDER BY qual_empno,qual_seq ASC"

Rs.Open Sql, Dbconn, 1

title_line = " 자격증 사항 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>자격 사항</title>
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
				<form action="insa_qual_view.asp?emp_no=<%=emp_no%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:60px; text-align:left" readonly="true">
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=emp_name%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="18%" >
							<col width="6%" >
                            <col width="10%" >
                            <col width="24%" >
                            <col width="*" >
                            <col width="16%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">자격증&nbsp;종목</th>
                                <th scope="col">등급</th>
                                <th scope="col">합격년월일</th>
                                <th scope="col">발급&nbsp;기관</th>
                                <th scope="col">자격증&nbsp;&nbsp;번호</th>
                                <th scope="col">경력수첩No.</th>
                                <th scope="col">자격수당</th>
 							</tr>
						</thead>
						<tbody>
						<%
						    v_cnt = 0
							do until rs.eof or rs.bof
							    v_cnt = v_cnt + 1
								if rs("qual_pay_id") = "" or isnull(rs("qual_pay_id")) then
								        qual_pay_id = "N"
								   else
								        qual_pay_id = rs("qual_pay_id")
								end if
						%>
							<tr>
								<td><%=rs("qual_type")%>&nbsp;</td>
								<td><%=rs("qual_grade")%>&nbsp;</td>
                                <td><%=rs("qual_pass_date")%>&nbsp;</td>
                                <td><%=rs("qual_org")%>&nbsp;</td>
                                <td><%=rs("qual_no")%>&nbsp;</td>
                                <td><%=rs("qual_passport")%>&nbsp;</td>
                                <td style=" border-bottom:1px solid #e3e3e3;"><%=qual_pay_id%>&nbsp;</td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
                            <% if v_cnt = 0 then %>
								     <td class="first" colspan="6">내역이 없습니다.</td>
                            <%    else %>
                                     <td class="first" colspan="6"><%=v_cnt%>&nbsp;건이 조회되었습니다.</td>
                            <% end if %>
							</tr>
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

