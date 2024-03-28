<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim u_type, org_code, curr_date, code_last, org_owner_date
Dim rsOrg, org_level, org_date, org_empno, org_empname
Dim org_company, org_bonbu, org_saupbu, org_team, org_reside_place, org_reside_company
Dim org_cost_group, org_cost_center, owner_org, owner_empno, owner_empname
Dim org_table_org, tel_ddd, tel_no1, tel_no2, org_sido, org_gugun, org_dong
Dim org_addr, org_end_date, org_reg_date, org_reg_user, org_mod_date, org_mod_user
Dim owner_orgname, title_line, rsCost, trade_code

Dim view_condi

u_type = Request.QueryString("u_type")
org_code = Request.QueryString("org_code")

curr_date = Mid(CStr(Now()), 1, 10)
code_last = ""
org_owner_date = curr_date

If u_type = "U" Then
	'Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	objBuilder.Append "SELECT org_level, org_name, org_date, org_end_date, org_empno, org_emp_name, "
	objBuilder.Append "	org_company, org_bonbu, org_saupbu, org_team, org_reside_place, org_reside_company, "
	objBuilder.Append "	org_cost_group, org_cost_center, org_owner_org, org_owner_empno, org_owner_empname, "
	objBuilder.Append "	org_table_org, org_tel_ddd, org_tel_no1, org_tel_no2, org_sido, org_gugun, org_dong, "
	objBuilder.Append "	org_addr, org_end_date, org_reg_date, org_reg_user, org_mod_date, org_mod_user, "
	objBuilder.Append "	(SELECT org_name FROM emp_org_mst WHERE org_code = eomt.org_owner_org) AS 'owner_orgname' "
	objBuilder.Append "FROM emp_org_mst AS eomt "
	objBuilder.Append "WHERE org_code = '"&org_code&"';"

	Set rsOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    org_level = rsOrg("org_level")
    org_name = rsOrg("org_name")
    org_date = rsOrg("org_date")
	org_end_date = rsOrg("org_end_date")
    org_empno = rsOrg("org_empno")
    org_empname = rsOrg("org_emp_name")
    org_company = rsOrg("org_company")
    org_bonbu = rsOrg("org_bonbu")
    org_saupbu = rsOrg("org_saupbu")
    org_team = rsOrg("org_team")
	org_reside_place = rsOrg("org_reside_place")
	org_reside_company = rsOrg("org_reside_company")
	org_cost_group = rsOrg("org_cost_group")
	org_cost_center = rsOrg("org_cost_center")
    owner_org = rsOrg("org_owner_org")
    owner_empno = rsOrg("org_owner_empno")
    owner_empname = rsOrg("org_owner_empname")

	If f_toString(rsOrg("org_table_org"), "") = "" Then
	        org_table_org = 0
	Else
			org_table_org = rsOrg("org_table_org")
	End If

    tel_ddd = rsOrg("org_tel_ddd")
    tel_no1 = rsOrg("org_tel_no1")
    tel_no2 = rsOrg("org_tel_no2")
	org_sido = rsOrg("org_sido")
    org_gugun = rsOrg("org_gugun")
    org_dong = rsOrg("org_dong")
    org_addr = rsOrg("org_addr")
    org_end_date = rsOrg("org_end_date")
    org_reg_date = rsOrg("org_reg_date")
	org_reg_user = rsOrg("org_reg_user")
    org_mod_date = rsOrg("org_mod_date")
    org_mod_user = rsOrg("org_mod_user")

	'org_owner_date = rs("org_owner_date")
	org_owner_date = curr_date

	owner_orgname = rsOrg("owner_orgname")

	rsOrg.close() : Set rsOrg = Nothing

	'Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
	'Set rs_owner=DbConn.Execute(Sql)

    'owner_orgname = rs_owner("org_name")
	'rs_owner.close()

	title_line = " 조직 상위조직 변경(발령) "
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사 관리 시스템</title>
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

		//조직 변경일
		$(function(){
			$( "#datepicker" ).datepicker();
			$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
			$( "#datepicker" ).datepicker("setDate", "<%=org_owner_date%>" );
		});

		//조직 폐쇄일
		$(function(){
			$( "#datepicker1" ).datepicker();
			$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
			$( "#datepicker1" ).datepicker("setDate", "<%=org_end_date%>" );
		});

		function goAction(){
		   window.close();
		}

		function goBefore(){
		   history.back();
		}

		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit ();
			}
		}

		function chkfrm(){
			if(document.frm.org_code.value == ""){
				alert('조직코드를 입력해주세요.');
				frm.org_code.focus();
				return false;
			}

			if(document.frm.org_name.value == ""){
				alert('조직명을 입력해주세요');
				frm.org_name.focus();
				return false;
			}

			if(document.frm.org_date.value == ""){
				alert('조직생성일을 입력해주세요');
				frm.org_date.focus();
				return false;
			}

			/*if(document.frm.org_empno.value == ""){
				alert('조직장사번을 입력해주세요');
				frm.org_empno.focus();
				return false;
			}

			if(document.frm.org_empname.value == ""){
				alert('조직장성명을 입력해주세요');
				frm.org_empname.focus();
				return false;
			}

			if(document.frm.org_cost_group.value =="") {
				alert('비용센타그룹을 선택하세요');
				frm.org_cost_group.focus();
				return false;
			}*/

			if(document.frm.org_cost_center.value == "상주직접비"){
				if(document.frm.org_reside_place.value == ""){
					alert('상주처를 입력해주세요');
					frm.org_reside_place.focus();
					return false;
				}
			}

			if(document.frm.org_cost_center.value =="상주직접비"){
				if(document.frm.org_reside_company.value ==""){
					alert('상주처회사를 입력해주세요');
					frm.org_reside_company.focus();
					return false;
				}
			}

			if(document.frm.org_level.value !="회사"){
				if(document.frm.owner_org.value ==""){
					alert('상위조직을 입력해주세요');
					frm.owner_org.focus();
					return false;
				}
			}

			if(document.frm.org_level.value =="상주처"){
				if(document.frm.org_reside_company.value ==""){
					alert('상주처 회사를 입력해주세요');
					frm.org_reside_company.focus();
					return false;
				}
			}

			if(document.frm.org_level.value =="상주처"){
				if(document.frm.org_cost_center.value !="상주직접비"){
					alert('비용구분에 상주직접비를 입력해주세요');
					frm.org_cost_center.focus();
					return false;
				}
			}

			var result = confirm('등록 하시겠습니까?');

			if(!result){
				return false;
			}
			return true;
		}

		function num_chk(txtObj){
			org_to = parseInt(document.frm.org_table_org.value.replace(/,/g,""));

			org_to = String(org_to);
			num_len = org_to.length;
			sil_len = num_len;
			org_to = String(org_to);
			if (org_to.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) org_to = org_to.substr(0,num_len -3) + "," + org_to.substr(num_len -3,3);
			if (sil_len > 6) org_to = org_to.substr(0,num_len -6) + "," + org_to.substr(num_len -6,3) + "," + org_to.substr(num_len -2,3);
			document.frm.org_table_org.value = org_to;
		}

		//주소 조회
		function jusoCallBack(roadFullAddr,roadAddrPart1,addrDetail,roadAddrPart2,engAddr,jibunAddr,zipNo,admCd,rnMgtSn,bdMgtSn,detBdNmList,bdNm,bdKdcd,siNm,sggNm,emdNm,liNm,rn,udrtYn,buldMnnm,buldSlno,mtYn,lnbrMnnm,lnbrSlno,emdNo,gubun){
				if(gubun === 'org'){
					$('#org_sido').val(siNm);
					$('#org_gugun').val(sggNm);
					$('#org_dong').val(rn + ' ' + buldMnnm);
					$('#org_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#org_zipcode').val(zipNo);
				}
			}
	</script>
	<style type="text/css">
		.no-input{
			color:gray;
			background-color:#E0E0E0;
			border:1px solid #999999;
		}
	</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
	<div id="wrap">

		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/insa/insa_org_owner_modify_save.asp" method="post" name="frm">
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableWrite">
					<colgroup>
						<col width="8%" >
						<col width="17%" >
						<col width="8%" >
						<col width="17%" >
						<col width="8%" >
						<col width="17%" >
						<col width="8%" >
						<col width="17%" >
					</colgroup>
					<tbody>
						<tr>
							<th class="first">조직코드</th>
							<td class="left">
								<input type="text" name="org_code" value="<%=org_code%>" class="no-input" readonly/>
							</td>
							<th>조직명</th>
							<td class="left">
								<input type="text" name="org_name" value="<%=org_name%>" class="no-input" readonly/>
							</td>
							<th>조직&nbsp;구분</th>
							<td class="left">
								<input type="text" name="org_level" value="<%=org_level%>" class="no-input" readonly/>
							</td>
							<th>조직생성일</th>
							<td class="left">
								<input type="text" name="org_date" size="10" value="<%=org_date%>" readonly/>
							</td>
						 </tr>
						 <tr>
							<th class="first">조직장사번</th>
							<td class="left">
								<input type="text" name="org_empno" id="org_empno" size="7" value="<%=org_empno%>" class="no-input" readonly/>
							</td>
							<th>조직장성명</th>
							<td class="left">
								<input type="text" name="org_empname" id="org_empname" size="10" value="<%=org_empname%>" class="no-input" readonly/>
							</td>
							<th>소속</th>
							<td colspan="3" class="left">
								<input type="text" name="org_company" id="org_company" style="width:100px;" value="<%=org_company%>" class="no-input" readonly/>
								<input type="text" name="org_bonbu" id="org_bonbu" style="width:100px;" value="<%=org_bonbu%>" class="no-input" readonly/>
								<input type="text" name="org_saupbu" id="org_saupbu" style="width:100px;" value="<%=org_saupbu%>" class="no-input" readonly/>
								<input type="text" name="org_team" id="org_team" style="width:100px;" value="<%=org_team%>" class="no-input" readonly/>
							</td>
						 </tr>
						<tr>
							<th class="first" style="background:#FFC">상위조직코드</th>
							<td class="left">
								<input type="text" name="owner_org" id="owner_org" size="4" value="<%=owner_org%>" readonly/>
								<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=owner2&mg_level=<%=org_level%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=850,height=400')">상위조직찾기</a>
							</td>
							<th>상위조직명</th>
							<td class="left">
								<input type="text" name="owner_orgname" id="owner_orgname" size="20" value="<%=owner_orgname%>" readonly/>
							</td>
							<th>상위조직장</th>
							<td class="left">
								<input type="text" name="owner_empno" id="owner_empno" size="7" value="<%=owner_empno%>" readonly/>
							</td>
							<th>상위조직장명</th>
							<td class="left">
								<input type="text" name="owner_empname" id="owner_empname" size="20" value="<%=owner_empname%>" readonly/>
							</td>
						 </tr>
						 <tr>
							<th class="first">대표전화</th>
							<td class="left">
								<input type="text" name="tel_ddd" id="tel_ddd" style="width:30px;" maxlength="3" value="<%=tel_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								-
								<input type="text" name="tel_no1" id="tel_no1" style="width:40px;" maxlength="4" value="<%=tel_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								-
								<input type="text" name="tel_no2" id="tel_no2" style="width:40px;" maxlength="4" value="<%=tel_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
							</td>
							<th>조직폐쇄일</th>
							<td class="left">
								<input type="text" name="org_end_date" size="10" id="datepicker1" style="width:70px;" value="<%=org_end_date%>" readonly/>
							</td>
							<th>상주처</th>
							<td class="left">
								<input type="text" name="org_reside_place" id="org_reside_place" style="width:150px;" value="<%=org_reside_place%>"/>
							</td>
							<th class="first">상주처 회사</th>
							<td class="left">
								<input type="text" name="org_reside_company" id="org_reside_company" style="width:120px" readonly="true" value="<%=org_reside_company%>"/>
								<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_trade_search.asp?gubun=1','tradesearch','scrollbars=yes,width=600,height=400')">찾기</a>
							</td>

						 </tr>
						 <tr>
							<th class="first">주소</th>
							<td colspan="5" class="left">
								<input type="text" name="org_sido" id="org_sido" style="width:100px" readonly="true" value="<%=org_sido%>">
								<input type="text" name="org_gugun" id="org_gugun" style="width:120px" readonly="true" value="<%=org_gugun%>">
								<input type="text" name="org_dong" id="org_dong" style="width:120px" readonly="true" value="<%=org_dong%>">
								<input type="text" name="org_addr" id="org_addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=org_addr%>">
								<input type="hidden" name="org_zip" id="org_zip" value="">

								<!--<a href="#" class="btnType03" onClick="pop_Window('/insa/zipcode_search.asp?gubun=org','org_zip_select','scrollbars=yes,width=600,height=400')">주소조회</a>-->
								<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=org','주소 조회','scrollbars=yes,width=600,height=400')">주소조회</a>
							</td>
							<th>비용센타그룹</th>
							<td class="left">
								<input type="text" name="org_cost_group" id="org_cost_group" style="width:150px" value="<%=org_cost_group%>" readonly/>
							</td>
						  </tr>
						  <tr>
							<th style="background:#FFC">조직변경일</th>
							<td class="left">
								<input type="text" name="org_owner_date" size="10" readonly="true" id="datepicker" style="width:70px;" value="<%=org_owner_date%>"/>
							</td>
							<th>적정인원(T.O)</th>
							<td colspan="3" class="left">
								<input type="text" name="org_table_org" id="org_table_org" style="width:90px;text-align:right" value="<%=FormatNumber(org_table_org, 0)%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
							</td>
							<th>비용구분</th>
							<td class="left">
						  <%
							objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
							objBuilder.Append "WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC;"

							Set rsCost = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()
						  %>
							<select name="org_cost_center" id="org_cost_center" style="width:90px;">
								<option value="" <%If org_cost_center = "" Then %>selected<%End If%>>선택</option>
						  <%
							Do Until rsCost.EOF
						  %>
								<option value='<%=rsCost("emp_etc_name")%>' <%If org_cost_center = rsCost("emp_etc_name") Then %>selected<%End If %>><%=rsCost("emp_etc_name")%></option>
						  <%
								rsCost.MoveNext()
							Loop
							rsCost.Close() : Set rsCost = Nothing
							DBConn.Close() : Set DBConn = Nothing
						  %>
							 </select>
							</td>
						  <tr>
							<th>등록일자</th>
							<td class="left">
								<input type="text" name="org_reg_date" id="org_reg_date" style="width:150px" value="<%=org_reg_date%>" class="no-input" readonly/>
							</td>
							<th>수정일자</th>
							<td class="left" colspan="5">
								<input type="text" name="org_mod_date" id="org_mod_date" style="width:150px" value="<%=org_mod_date%>" class="no-input" readonly/>
							</td>
						  </tr>
					</tbody>
				</table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01">
					<input type="button" value="저장" onclick="javascript:frmcheck();"/>
				</span>
				<%If u_type = "U" Then%>
					  <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
				<%Else%>
					 <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"/></span>
				<%End If%>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			<input type="hidden" name="mg_level" value="<%=org_level%>"/>
			<input type="hidden" name="trade_code" value="<%=trade_code%>"/>
		</form>
	</div>
</div>
</body>
</html>