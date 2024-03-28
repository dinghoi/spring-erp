<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim title_line, rsQual

title_line = "자격 사항"

If m_seq = "" Or m_name = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('회원기본가입 등록 후 이용 가능합니다.');"
	Response.Write "	location.href='/member/member_add.asp';"
	Response.Write "</script>"

	Response.End
End If

objBuilder.Append "SELECT qual_type,qual_grade, qual_pass_date, qual_org, qual_no, "
objBuilder.Append "	qual_passport, qual_seq "
objBuilder.Append "FROM member_qual "
objBuilder.Append "WHERE m_seq = '"&m_seq&"' "
objBuilder.Append "ORDER BY m_seq, qual_seq ASC;"

Set rsQual = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회원 관리</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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

			//자격사항 등록 팝업
			function qualAddPopup(){
				var url = '/member/member_qual_add.asp';
				var pop_name = '학력사항 등록';
				//var param = '?m_seq='+id+'&m_name='+name;
				var features = 'scrollbars=yes,width=750,height=300';

				//url += param;

				pop_Window(url, pop_name, features);
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
                            <strong>성명 : </strong>
							<label>
								<input type="text" name="m_name" id="m_name" value="<%=m_name%>" readonly="true" style="width:150px; text-align:left"/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="8%" >
							<col width="9%" >
							<col width="15%" >
							<col width="*" >
							<col width="15%" >
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                            <th>자격증 종목</th>
                            <th>등급</th>
                            <th>합격년월일</th>
                            <th>발급 기관명</th>
                            <th>자격 등록번호</th>
                            <th>경력수첩No.</th>
                            <th>순번</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsQual.EOF Or rsQual.BOF Then
							Response.Write "<tr><td colspan='7' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsQual.EOF
						%>
							<tr>
								<td><%=rsQual("qual_type")%>&nbsp;</td>
								<td><%=rsQual("qual_grade")%>&nbsp;</td>
								<td><%=rsQual("qual_pass_date")%>&nbsp;</td>
								<td><%=rsQual("qual_org")%>&nbsp;</td>
								<td><%=rsQual("qual_no")%>&nbsp;</td>
								<td><%=rsQual("qual_passport")%>&nbsp;</td>
								<td class="right"><%=rsQual("qual_seq")%>&nbsp;</td>
							</tr>
						<%
								rsQual.MoveNext()
							Loop
						End If
						rsQual.Close() : Set rsQual = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="qualAddPopup();" class="btnType04">자격사항 등록</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>