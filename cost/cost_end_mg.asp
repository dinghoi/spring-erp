<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim rs_insa, max_org_month
Dim org_bonbu, rsOrgCode
Dim rsEmpOrgList
Dim title_line

objBuilder.Append "SELECT MAX(org_month) AS max_org_month "
objBuilder.Append "FROM emp_org_mst_month "

Set rs_insa = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_insa("max_org_month")) Then
    max_org_month = "000000"
Else
    max_org_month = rs_insa("max_org_month")
End If
rs_insa.Close() : Set rs_cost = Nothing

'비용 권한이 0이 아닐 경우 본부명 검색[허정호_20210306]
If cost_grade <> "0" Then
	objBuilder.Append "SELECT eomt.org_bonbu "
	objBuilder.Append "FROM emp_master_month AS emmt "
	objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
	objBuilder.Append "WHERE emmt.emp_month = '"&max_org_month&"' "
	objBuilder.Append "	AND emmt.emp_no = '"&emp_no&"' "

	Set rsOrgCode = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	org_bonbu = rsOrgCode("org_bonbu")

	rsOrgCode.Close() : Set rsOrgCode = Nothing
End If

' 2019.02.22 박정신 요구 'N/W 1사업부','N/W 2사업부',"SI3사업부","솔루션사업부"	는 나오지않도록 조건으로 처리..
'sql = "SELECT *                                " & chr(13) & _
'      "  FROM emp_org_mst                      " & chr(13) & _
'      " WHERE (org_level = '사업부')           " & chr(13) & _
'      "   AND (org_name <> '총괄대표')         " & chr(13) & _
'      "   AND (    ISNULL(org_end_date)        " & chr(13) & _
'      "         OR org_end_date = '0000-00-00' " & chr(13) & _
'      "       )                                " & chr(13)
' org_end_date = '' or   ....   date형을 '' 으로 비교할수없다.   Warning: Incorrect date value: '' for column 'org_end_date' at row 1

objBuilder.Append "SELECT org_name, org_date "
objBuilder.Append "FROM emp_org_mst "
objBuilder.Append "WHERE org_level = '본부' "
objBuilder.Append "	AND (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
'objBuilder.Append "	AND org_name NOT IN ('전략부문', 'ICT연구소', '빅데이타연구소', '기술연구소', '한진그룹사업본부')"
objBuilder.Append "	AND org_name NOT IN ('전략부문', 'ICT연구소', '기술연구소', '한진그룹사업본부', 'SI수행본부', '스마트본부')"

If cost_grade = "0" Then
	objBuilder.Append "GROUP BY org_bonbu, org_name "
	objBuilder.Append "ORDER BY FIELD(org_company, '케이원', '케이네트웍스', '케이시스템'), "
	objBuilder.Append "	FIELD(org_bonbu, '빅데이타연구소', '스마트본부', 'DI사업부문', '공공SI본부', '금융SI본부', 'ICT본부', '공공본부', 'NI본부', 'SI2본부', 'SI1본부') DESC "
Else
	objBuilder.Append "	AND (org_name = '"&org_bonbu&"' Or org_empno = '"&emp_no&"') "
	objBuilder.Append "GROUP BY org_name "
End If

Set rsEmpOrgList = DBConn.Execute(objBuilder.ToSTring)
objBuilder.Clear()

title_line = "비용 마감 관리"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	    <script src="/java/jquery-1.9.1.js"></script>
	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}

			function frmcheck(){
				document.frm.submit();
			}

			//비용마감 처리
			function setCostEnd(end_url, end_month, dept, end_yn, type){
				if(type == 'A'){
					if(!confirm('해당 월(' +end_month + ')의 비용 마감을 일괄 진행하시겠습니까?')){
						return false;
					}
				}

				var param = {"end_month":end_month, "saupbu":encodeURIComponent(dept), "end_yn":end_yn};

				let start_time = new Date();

				$.ajax({
					type : "GET"
					, dataType : 'html'
					, contentType: "application/x-www-form-urlencoded; charset=EUC-KR"
					, url: end_url
					, data: param
					, async: true
					, error: function(request, status, error){
						console.log("code = "+ request.status + " message = " + request.responseText + " error = " + error);
					}
					, success: function(data){
						let end_time = new Date();
						var elapedMin = (end_time.getTime() - start_time.getTime()) / 1000 / 60;

						console.log('진행시간(분) : ' + elapedMin);
						console.log(data);
						console.log($(window).scrollTop());

						alert(data);
						location.href="/cost/cost_end_mg.asp";
						return;
					}
					, beforeSend: function(){
						var width = 0;
						var height = 0;
						var left = 0;
						var top = 0;

						width = 220;
						height = 118;
						top = ( $(window).height() - height ) / 2 + $(window).scrollTop();
						left = ( $(window).width() - width ) / 2 + $(window).scrollLeft();

						if($("#div_ajax_load_image").length != 0){
							$("#div_ajax_load_image").css({
								"top": top+"px",
								"left": left+"px"
							});
							$("#div_ajax_load_image").show();
						}else{
							$('body').append('<div id="div_ajax_load_image" style="position:absolute; top:' + top + 'px; left:' + left + 'px; width:' + width + 'px; height:' + height + 'px; z-index:9999; background:#f0f0f0; filter:alpha(opacity=50); opacity:alpha*0.5; margin:auto; padding:0; "><img src="/image/wait.gif" style="width:220px; height:118px;"></div>');
						}
					}
					, complete: function(){
						$("#div_ajax_load_image").hide();
					}
				});
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/cost_end_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
						<dd>
						<p>
							<label>&nbsp;&nbsp;<strong>최신정보로 다시 조회하기&nbsp;</strong></label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
							(영업부 월마감,인사부 조직코드마감, 인사마감[<%=max_org_month%>] 확인)
						</p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="13%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">사업부</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">현 재 마 감 현 황</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">새로운 마감 처리</th>
								<th rowspan="2" scope="col">보고자료</th>
								<th rowspan="2" scope="col">본부장보고</th>
								<th rowspan="2" scope="col">CEO보고</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">마감년월</th>
							  <th scope="col">마감상태</th>
							  <th scope="col">마감자</th>
							  <th scope="col">처리일자</th>
							  <th scope="col">마감취소</th>
							  <th scope="col">마감년월</th>
							  <th scope="col">마감처리</th>
              				</tr>
						</thead>
						<tbody>
						<%
							Dim rsCostEndMax, rsCostEndList

							Dim cancel_yn, rs_cost, new_date
							Dim new_month, now_month, end_view, end_yn, end_month
							Dim reg_name, reg_id, reg_date, batch_view, bonbu_view
							Dim ceo_view, batch_yn, bonbu_yn, ceo_yn

							Dim jik_yn

							'=====	사업부 별 마감 항목 리스트	=====
							Do Until rsEmpOrgList.EOF
								cancel_yn = "N"

								'If rs("org_bonbu") = "직할사업부" Then
								'	If rs("org_saupbu") = "공항지원사업부" Or rs("org_saupbu") = "KAL지원사업부" Then
								'		jik_yn = "N"
								'	Else
								'		jik_yn = "Y"
							  	'	End If
								'Else
							  	'	jik_yn = "N"
								'End If

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '"&rsEmpOrgList("org_name")&"' "

								Set rsCostEndMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '"&rsEmpOrgList("org_name")&"' "
								objBuilder.Append "	AND end_month = '"&rsCostEndMax("max_month")&"' "

								rsCostEndMax.Close()

								Set rsCostEndList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndList.EOF Or rsCostEndList.BOF Then
									new_date = DateAdd("m", -1, Now())
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "없음"

									If end_month = "없음" Then
										new_date = rsEmpOrgList("org_date")
										new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									End If

									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								 Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndList("end_month"), 1, 4) & "-" & Mid(rsCostEndList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndList("end_yn") = "Y" Then
										end_view = "마감"
									ElseIf rsCostEndList("end_yn") = "C" Then
										new_month = rsCostEndList("end_month")
										end_view = "취소"
									Else
										end_view = "진행"
									End If

									end_yn = rsCostEndList("end_yn")
									end_month = rsCostEndList("end_month")
									reg_name = rsCostEndList("reg_name")
									reg_id = rsCostEndList("reg_id")
									reg_date = rsCostEndList("reg_date")

									If rsCostEndList("batch_yn") = "Y" Then
										batch_view = "자료생성"
									Else
								  		batch_view = "미생성"
									End If

									If rsCostEndList("bonbu_yn") = "Y" Then
										bonbu_view = "승인완료"
									End If

									If rsCostEndList("ceo_yn") = "Y" Then
										ceo_view = "승인완료"
									End If

									If rsCostEndList("batch_yn") = "Y" And rsCostEndList("bonbu_yn") = "N" Then
										bonbu_view = "진행중"
									  	ceo_view = ""
									End If

									If rsCostEndList("bonbu_yn") = "Y" And rsCostEndList("ceo_yn") = "N" Then
										ceo_view = "진행중"
									End If

									If rsCostEndList("batch_yn") = "N" And rsCostEndList("bonbu_yn") = "N" And rsCostEndList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndList("batch_yn")
									bonbu_yn = rsCostEndList("bonbu_yn")
									ceo_yn = rsCostEndList("ceo_yn")
								End If

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
							  		If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr>
								<td class="first"><%=rsEmpOrgList("org_name")%></td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "취소" Then
										Response.Write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.Write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									If cancel_yn = "Y" Then
										Response.write "<a href='/cost/cost_end_cancel.asp?saupbu="&rsEmpOrgList("org_name")&"&end_month="&end_month&"' class='btnType03'>마감취소</a>"
									Else
										Response.write "취소불가"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>"
									<%If SysAdminYn <> "Y" Then%> readonly="true" <%End If%> />
								</td>
								<td>
									<%
									if now_month > new_month then
										'Response.write "<a href='/cost/cost_end_pro.asp?saupbu="&rsEmpOrgList("org_name")&"&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/cost_end_pro.asp"", """&new_month&""", """&rsEmpOrgList("org_name")&""", """&end_yn&""", ""S"");' class='btnType03'>마감</a>"
									else
										Response.write "마감불가"
									end if
									%>
                				</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
							<%
								rsEmpOrgList.MoveNext()
							Loop
							rsCostEndList.Close() : Set rsCostEndList = Nothing
							Set rsCostEndMax = Nothing
							rsEmpOrgList.Close() : Set rsEmpOrgList = Nothing

							'=====	사업부외나머지	=====
							Dim rsCostEndEtcList, rsCostEndEtcMax

							If cost_grade = "0" Then
								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu='사업부외나머지' "

								Set rsCostEndEtcMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '사업부외나머지' "
								objBuilder.Append "	AND end_month = '"&rsCostEndEtcMax("max_month")&"' "

								rsCostEndEtcMax.Close() : Set rsCostEndEtcMax = Nothing

								Set rsCostEndEtcList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndEtcList.EOF Or rsCostEndEtcList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "없음"
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndEtcList("end_month"), 1, 4) & "-" & Mid(rsCostEndEtcList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndEtcList("end_yn") = "Y" Then
										end_view = "마감"
									ElseIf rsCostEndEtcList("end_yn") = "C" Then
										new_month = rsCostEndEtcList("end_month")
										end_view = "취소"
									Else
										end_view = "진행"
									End If

									end_yn = rsCostEndEtcList("end_yn")
									end_month = rsCostEndEtcList("end_month")
									reg_name = rsCostEndEtcList("reg_name")
									reg_id = rsCostEndEtcList("reg_id")
									reg_date = rsCostEndEtcList("reg_date")

									If rsCostEndEtcList("batch_yn") = "Y" Then
										batch_view = "자료생성"
									Else
										batch_view = "미생성"
									End If

									If rsCostEndEtcList("bonbu_yn") = "Y" Then
										bonbu_view = "승인완료"
									End If

									If rsCostEndEtcList("ceo_yn") = "Y" Then
										ceo_view = "승인완료"
									End If

									If rsCostEndEtcList("batch_yn") = "Y" And rsCostEndEtcList("bonbu_yn") = "N" Then
										bonbu_view = "진행중"
									  ceo_view = ""
									End If

									If rsCostEndEtcList("bonbu_yn") = "Y" And rsCostEndEtcList("ceo_yn") = "N" Then
										ceo_view = "진행중"
									End If

									If rsCostEndEtcList("batch_yn") = "N" And rsCostEndEtcList("bonbu_yn") = "N" And rsCostEndEtcList("ceo_yn") = "N" Then
										bonbu_view = ""
									  ceo_view = ""
									End If

									batch_yn = rsCostEndEtcList("batch_yn")
									bonbu_yn = rsCostEndEtcList("bonbu_yn")
									ceo_yn = rsCostEndEtcList("ceo_yn")
								End If
								rsCostEndEtcList.Close() : Set rsCostEndEtcList = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#FFE8E8">
								<td class="first">사업부외나머지</td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "취소" Then
										Response.Write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.Write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									If cancel_yn = "Y" Then
										Response.write "<a href='/cost/cost_bonbu_end_cancel.asp?saupbu=사업부외나머지&end_month="&end_month&"' class='btnType03'>마감취소</a>"
									Else
										Response.write "취소불가"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.write "<a href='/cost/cost_bonbu_end_pro.asp?saupbu=사업부외나머지&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/cost_bonbu_end_pro.asp"", """&new_month&""", ""사업부외나머지"", """&end_yn&""", ""S"");' class='btnType03'>마감</a>"
									Else
										Response.write "마감불가"
									End If
									%>
								</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
							<%
							End If

							'=====	상주 비용	=====
							If resideEndViewYn = "Y" Then
								Dim rsCostEndMonthMax, rsCostEndResideList

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '상주비용' "

								Set rsCostEndMonthMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end  "
								objBuilder.Append "WHERE saupbu = '상주비용' "
								objBuilder.Append "	AND end_month = '"&rsCostEndMonthMax("max_month")&"'"

								Set rsCostEndResideList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndResideList.EOF Or rsCostEndResideList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "없음"
									end_yn = ""
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndResideList("end_month"), 1, 4) & "-" & Mid(rsCostEndResideList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndResideList("end_yn") = "Y" Then
										end_view = "마감"
									ElseIf rsCostEndResideList("end_yn") = "C" Then
										new_month = rsCostEndResideList("end_month")
										end_view = "취소"
									Else
										end_view = "진행"
									End If

									end_yn = rsCostEndResideList("end_yn")
									end_month = rsCostEndResideList("end_month")
									reg_name = rsCostEndResideList("reg_name")
									reg_id = rsCostEndResideList("reg_id")
									reg_date = rsCostEndResideList("reg_date")

									If rsCostEndResideList("batch_yn") = "Y" Then
										batch_view = "자료생성"
									Else
										batch_view = "미생성"
									End If

									If rsCostEndResideList("bonbu_yn") = "Y" Then
										bonbu_view = "승인완료"
									End If

									If rsCostEndResideList("ceo_yn") = "Y" Then
										ceo_view = "승인완료"
									End If

									If rsCostEndResideList("batch_yn") = "Y" And rsCostEndResideList("bonbu_yn") = "N" Then
										bonbu_view = "진행중"
									  ceo_view = ""
									End If

									If rsCostEndResideList("bonbu_yn") = "Y" And rsCostEndResideList("ceo_yn") = "N" Then
										ceo_view = "진행중"
									End If

									If rsCostEndResideList("batch_yn") = "N" And rsCostEndResideList("bonbu_yn") = "N" And rsCostEndResideList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndResideList("batch_yn")
									bonbu_yn = rsCostEndResideList("bonbu_yn")
									ceo_yn = rsCostEndResideList("ceo_yn")
								End If

								rsCostEndResideList.Close() : Set rsCostEndResideList = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#FFFFCC">
								<td class="first">상주비용</td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "취소" Then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.write end_view
									End If
									%>&nbsp;
                				</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  <td>
									<%
									If cancel_yn = "Y" Then
										Response.Write "<a href='/cost/company_cost_end_cancel.asp?end_month="&end_month&"'  class='btnType03'>마감취소</a>"
									Else
										Response.Write "취소불가"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.Write "<a href='/cost/company_cost_end_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/company_cost_end_pro.asp"", """&new_month&""", """", """&end_yn&""", ""S"");' class='btnType03'>마감</a>"
									Else
										Response.Write "마감불가"
									End If
									%>
								</td>
							  	<td colspan="3">&nbsp;</td>
							</tr>
								<%'=====	공통비/직접비배분		=====
								Dim rsCostEndCommList, rsCostEndCommMax

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu='공통비/직접비배분' "

								Set rsCostEndCommMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '공통비/직접비배분' "
								objBuilder.Append "	AND end_month ='"&rsCostEndCommMax("max_month")&"' "

								rsCostEndCommMax.Close() : Set rsCostEndCommMax = Nothing

								Set rsCostEndCommList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndCommList.EOF Or rsCostEndCommList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "없음"
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndCommList("end_month"), 1, 4) & "-" & Mid(rsCostEndCommList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndCommList("end_yn") = "Y" Then
										end_view = "마감"
									ElseIf rsCostEndCommList("end_yn") = "C" Then
										new_month = rsCostEndCommList("end_month")
										end_view = "취소"
									Else
										end_view = "진행"
									End If

									end_yn = rsCostEndCommList("end_yn")
									end_month = rsCostEndCommList("end_month")
									reg_name = rsCostEndCommList("reg_name")
									reg_id = rsCostEndCommList("reg_id")
									reg_date = rsCostEndCommList("reg_date")

									If rsCostEndCommList("batch_yn") = "Y" Then
										batch_view = "자료생성"
									Else
										batch_view = "미생성"
									End If

									If rsCostEndCommList("bonbu_yn") = "Y" Then
										bonbu_view = "승인완료"
									End If

									If rsCostEndCommList("ceo_yn") = "Y" Then
										ceo_view = "승인완료"
									End If

									If rsCostEndCommList("batch_yn") = "Y" And rsCostEndCommList("bonbu_yn") = "N" Then
										bonbu_view = "진행중"
										ceo_view = ""
									End If

									If rsCostEndCommList("bonbu_yn") = "Y" And rsCostEndCommList("ceo_yn") = "N" Then
									  ceo_view = "진행중"
									End If

									If rsCostEndCommList("batch_yn") = "N" And rsCostEndCommList("bonbu_yn") = "N" And rsCostEndCommList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndCommList("batch_yn")
									bonbu_yn = rsCostEndCommList("bonbu_yn")
									ceo_yn = rsCostEndCommList("ceo_yn")
								End If

								rsCostEndCommList.Close() : Set rsCostEndCommList = Nothing
								DBConn.Close() : Set DBConn = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#CCFFFF">
								<td class="first">공통비/직접비배분</td>
					  	  		<td><%=end_month%></td>
								<td>
									<%
									If end_view = "취소" Then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							  		<%
							  		If cancel_yn = "Y" Then
							  			Response.Write "<a href='/cost/company_as_sum_cancel.asp?end_month="&end_month&"' class='btnType03'>마감취소</a>"
									Else
										Response.Write "취소불가"
									End If
									%>
                				</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.Write "<a href='/cost/company_as_sum_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/company_as_sum_pro.asp"", """&new_month&""", """", """&end_yn&""", ""S"");' class='btnType03'>마감</a>"
									Else
										Response.Write "마감불가"
									End If
									%>
								</td>
								<td colspan="3">&nbsp;</td>
						  	</tr>
							<%
							End If
							%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>