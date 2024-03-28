<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
gubun = Request("gubun")

draft_man = Request.form("draft_man")
if gubun = "" or isnull(gubun) then
	gubun = Request.form("gubun")
end if
'response.write(gubun)
Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect


   if draft_man = "" then
	   SQL = "select * from emp_alba_mst where draft_man = '" + draft_man + "' ORDER BY draft_man ASC"
    else
	   SQL = "select * from emp_alba_mst where draft_man like '%" + draft_man + "%' ORDER BY draft_man ASC"
   end if
   Rs.open SQL, Dbconn, 1


title_line = "아르바이트 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>아르바이트 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function alba_list(draft_no,draft_man,draft_tax_id,draft_date,company,bonbu,saupbu,team,org_name,sign_no,bank_name,account_no,account_name)
			{
				if(document.frm.gubun.value =="1") {
					opener.document.frm.draft_no.value = draft_no;
					opener.document.frm.draft_man.value = draft_man;
					opener.document.frm.draft_tax_id.value = draft_tax_id;
					opener.document.frm.company.value = company;
					opener.document.frm.bonbu.value = bonbu;
					opener.document.frm.saupbu.value = saupbu;
					opener.document.frm.team.value = team;
					opener.document.frm.org_name.value = org_name;
					opener.document.frm.sign_no.value = sign_no;
					opener.document.frm.bank_name.value = bank_name;
					opener.document.frm.account_no.value = account_no;
					opener.document.frm.account_name.value = account_name;
					window.close();
				}
				if(document.frm.gubun.value =="2") {
					opener.document.frm.cost_company.value = trade_name;
					window.close();
				}

			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.draft_man.value =="") {
					alert('아르바이트명을 입력하세요');
					frm.draft_man.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_alba_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>아르바이트명을 입력하세요 </strong>
								<label>
        						<input name="draft_man" type="text" id="draft_man" value="<%=draft_man%>" style="width:150px;text-align:left;ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
                            <col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">성명</th>
                                <th scope="col">주민번호</th>
								<th scope="col">소득구분</th>
								<th scope="col">업무등록일</th>
								<th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							draft_no = rs("draft_no")
							draft_man = rs("draft_man")
							person_no1 = rs("person_no1")
							person_no2 = rs("person_no2")
							draft_tax_id = rs("draft_tax_id")
							draft_date = rs("draft_date")
							work_memo = rs("work_memo")
							company = rs("company")
							bonbu = rs("bonbu")
							saupbu = rs("saupbu")
							team = rs("team")
							org_name = rs("org_name")
							sign_no = rs("sign_no")
							bank_name = rs("bank_name")
							account_no = rs("account_no")
							account_name = rs("account_name")
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="alba_list('<%=draft_no%>','<%=draft_man%>','<%=draft_tax_id%>','<%=draft_date%>','<%=company%>','<%=bonbu%>','<%=saupbu%>','<%=team%>','<%=org_name%>','<%=sign_no%>','<%=bank_name%>','<%=account_no%>','<%=account_name%>');"><%=rs("draft_man")%>(<%=rs("draft_no")%>)</a>
                                </td>
								<td><%=person_no1%>-<%=person_no2%>&nbsp;</td>
                                <td><%=draft_tax_id%>&nbsp;</td>
								<td><%=draft_date%>&nbsp;</td>
								<td><%=work_memo%>&nbsp;</td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="5">내역이 없습니다</td>
							</tr>
                            <tr>
								<td colspan="5">
					            <div class="btnRight">
					            <a href="#" onClick="pop_Window('insa_alba_add.asp?view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&u_type=<%=""%>','insa_alba_add_pop','scrollbars=yes,width=750,height=450')" class="btnType04">아르바이트 등록</a>
					           </div>  
                               </td>
							</tr>
                            
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			</form>
		</div>        				
	</body>
</html>

