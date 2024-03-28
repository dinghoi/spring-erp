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

Sql = "SELECT count(*) FROM emp_appoint where app_empno = '"&emp_no&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

if tottal_record = 0 then
   first_view = "Y"
end if

sql = "select * from emp_appoint where app_empno = '" + emp_no + "' ORDER BY app_empno,app_date,app_seq ASC"
'Response.write sql
Rs.Open Sql, Dbconn, 1

title_line = "◈ 발령 사항 ◈"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>발령 사항</title>
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
				<form action="insa_appoint_view.asp?emp_no=<%=emp_no%>" method="post" name="frm">
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
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="9%" >
							<col width="10%" >
							<col width="9%" >
							<col width="9%" >
							<col width="10%" >
                            <col width="9%" >
                            <col width="*" >
						</colgroup>
						<thead>
                            <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령일</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령구분</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령유형</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령전</th>
				                <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령후</th>
			                </tr>
                            <tr>
                                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">발령내용</th>
                            </tr>
						</thead>
						<tbody>
						<%
						     v_cnt = 0
							do until rs.eof or rs.bof
							 v_cnt = v_cnt + 1
						%>
							<tr>
								<td><%=rs("app_date")%>&nbsp;</td>
								<td><%=rs("app_id")%>&nbsp;</td>
                                <td><%=rs("app_id_type")%>&nbsp;</td>
                                <td><%=rs("app_to_company")%>&nbsp;</td>
                                <td><%=rs("app_to_orgcode")%>)<%=rs("app_to_org")%>&nbsp;</td>
                                <td><%=rs("app_to_grade")%>-<%=rs("app_to_position")%>&nbsp;</td>
                                <td><%=rs("app_be_company")%>&nbsp;</td>
                                <td><%=rs("app_be_orgcode")%>)<%=rs("app_be_org")%>&nbsp;</td>
                                <td><%=rs("app_be_grade")%>-<%=rs("app_be_position")%>&nbsp;</td>
                                <td class="left"><%=rs("app_start_date")%>&nbsp;-&nbsp;<%=rs("app_finish_date")%>&nbsp;<%=rs("app_be_enddate")%>&nbsp;<%=rs("app_reward")%>&nbsp;:&nbsp;<%=rs("app_comment")%>&nbsp;</td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						  if v_cnt = 0 then
						%>
							<tr>
								<td class="first" colspan="6">내역이 없습니다</td>
							</tr>
                        <%
						  end if
						%>

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

