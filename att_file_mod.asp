<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
acpt_no = request("acpt_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)
if rs.eof or rs.bof then
	acpt_user = "NO DATA"
  else
  	acpt_user = rs("acpt_user")
end if
rs.close()

Sql = "select * from att_file where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)

path = "/att_file/" + rs("company")

title_line = "첨부파일 변경 및 삭제"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=visit_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {

				if(document.frm.att_file1.value =="" && document.frm.att_file2.value =="" && document.frm.att_file3.value =="" && document.frm.att_file4.value =="" && document.frm.att_file5.value =="") {
					alert('사진 첨부가 되지 않았습니다');
					frm.att_file1.focus();
					return false;}

				{
				a=confirm('첨부파일을 변경 하시겠습니까?')
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
				<h3 class="tit"><%=title_line%></h3>
				<form action="att_file_mod_ok.asp" method="post" name="frm" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">고객명/접수번호</th>
								<td class="left"><%=acpt_user%>&nbsp;(<%=acpt_no%>)
                                <input name="acpt_no" type="hidden" id="acpt_no" value="<%=acpt_no%>">
								</td>
								<th>처리유형</th>
								<td class="left"><%=rs("as_type")%></td>
                                </td>
							</tr>
							<tr>
								<th class="first">회사</th>
								<td class="left"><%=rs("company")%></td>
								<th>부서</th>
								<td class="left"><%=rs("dept")%></td>              
							</tr>
						</tbody>
					</table>
			<h3 class="stit">* 첨부파일</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="40%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">기존 첨부파일</th>
								<th scope="col">변경 첨부파일</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first">1</td>
								<td>&nbsp;
							<% if rs("att_file1") <> "" then	%>		
								<a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file1")%>"><%=rs("att_file1")%></a>
                            <% end if	%>
                                </td>
								<td class="left"><input name="att_file1" type="file" id="att_file1" size="50"></td>
							</tr>
							<tr>
								<td class="first">2</td>
								<td>&nbsp;
							<% if rs("att_file2") <> "" then	%>		
								<a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file2")%>"><%=rs("att_file2")%></a>
                            <% end if	%>
                                </td>
								<td class="left"><input name="att_file2" type="file" id="att_file2" size="50"></td>
							</tr>
							<tr>
								<td class="first">3</td>
								<td>&nbsp;
							<% if rs("att_file3") <> "" then	%>		
								<a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file3")%>"><%=rs("att_file3")%></a>
                            <% end if	%>
                                </td>
								<td class="left"><input name="att_file3" type="file" id="att_file3" size="50"></td>
							</tr>
							<tr>
								<td class="first">4</td>
								<td>&nbsp;
							<% if rs("att_file4") <> "" then	%>		
								<a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file4")%>"><%=rs("att_file4")%></a>
                            <% end if	%>
                                </td>
								<td class="left"><input name="att_file4" type="file" id="att_file4" size="50"></td>
							</tr>
							<tr>
								<td class="first">5</td>
								<td>&nbsp;
							<% if rs("att_file5") <> "" then	%>		
								<a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file5")%>"><%=rs("att_file5")%></a>
                            <% end if	%>
                                </td>
								<td class="left"><input name="att_file5" type="file" id="att_file5" size="50"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
            	<input type="hidden" name="sido" value="<%=rs("sido")%>" ID="Hidden1">
                <input type="hidden" name="company" value="<%=rs("company")%>" ID="Hidden1">
                <input type="hidden" name="visit_date" value="<%=rs("visit_date")%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

