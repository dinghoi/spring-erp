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
Dim rs_emp, rsLng, title_line

If m_seq = "" Or m_name = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('회원기본가입 등록 후 이용 가능합니다.');"
	Response.Write "	location.href='/member/member_add.asp';"
	Response.Write "</script>"

	Response.End
End If

objBuilder.Append "SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date, "
objBuilder.Append "	lang_seq "
objBuilder.Append "FROM member_language "
objBuilder.Append "WHERE m_seq='"&m_seq&"' ORDER BY m_seq, lang_seq ASC"

Set rsLng = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "어학능력 사항"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회원 관리</title>
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

			//어학능력사항 등록 팝업
			function langAddPopup(){
				var url = '/member/member_language_add.asp';
				var pop_name = '어학능력사항 등록';
				var features = 'scrollbars=yes,width=750,height=300';

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
								<input name="m_name" type="text" id="m_name" value="<%=m_name%>" readonly="true" style="width:150px; text-align:left"/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="3">어학구분</th>
                                <th colspan="2">어학종류</th>
                                <th colspan="2">점수</th>
                                <th colspan="2">급수</th>
                                <th colspan="3">취득일</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsLng.EOF Or rsLng.BOF Then
							Response.Write "<tr><td colspan='12' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsLng.EOF
							%>
								<tr>
									<td colspan="3"><%=rsLng("lang_id")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_id_type")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_point")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_grade")%>&nbsp;</td>
									<td colspan="3"><%=rsLng("lang_get_date")%>&nbsp;</td>
								</tr>
							<%
								rsLng.MoveNext()
							Loop
						End If
						rsLng.Close() : Set rsLng = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="langAddPopup();" class="btnType04">어학사항 등록</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>