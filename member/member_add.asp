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
Dim be_pg, rsIndi, title_line
Dim rsEtc, rsMiliId, rsMiliGrade

If m_seq <> "" And m_name <> "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('회원가입 외 나머지 사항을 등록해주세요.');"
	Response.Write "	location.href='/member/member_family.asp';"
	Response.Write "</script>"

	Response.End
End If

be_pg = "/member/member_add_proc.asp"
title_line = "회원 기본정보"
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

			//생년월일
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
			});

			//결혼기념일
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "" );
			});

			//복무기간 시작일
			$(function(){
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "" );
			});

			//복부기간 종료일
			$(function(){
				$( "#datepicker3" ).datepicker();
				$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker3" ).datepicker("setDate", "" );
			});

			function goBefore(){
				var result = confirm("등록을 취소하시겠습니까?");

				if(result){
					location.href="/index.asp";
				}else{
					return false;
				}
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(isEmpty($('#m_name').val())){
					alert('성명(한글)을 입력해주세요');
					$('#m_name').focus();
					return false;
				}

				if(isEmpty($('#m_ename').val())){
					alert('성명(영문)을 입력해주세요');
					$('#m_ename').focus();
					return false;
				}

				if(isEmpty($('#datepicker').val())){
					alert('생년월일을 입력해주세요');
					$('#datepicker').focus();
					return false;
				}

				if(isEmpty($('#m_person1').val())){
					alert('주민번호를 입력해주세요');
					$('#m_person1').focus();
					return false;
				}

				if(isEmpty($('#m_person2').val())){
					alert('주민등록번호를 입력해주세요');
					frm.emp_person2.focus();
					return false;
				}

				/*if(isEmpty($('#m_sex').val())){
					alert('성별을 선택하세요');
					$('#m_sex').focus();
					return false;
				}*/

				if(isEmpty($('#m_hp_ddd').val())){
					alert('휴대폰번호를 입력해주세요');
					$('#m_hp_ddd').focus();
					return false;
				}

				if(isEmpty($('#m_hp_no1').val())){
					alert('휴대폰번호를 입력해주세요');
					$('#m_hp_no1').focus();
					return false;
				}

				if(isEmpty($('#m_hp_no2').val())){
					alert('휴대폰번호를 입력해주세요');
					$('#m_hp_no2').focus()
					return false;
				}

				if(isEmpty($('#m_emergency_tel').val())){
					alert('비상연락번호를 입력해주세요');
					$('#m_emergency_tel').focus();
					return false;
				}

				if(isEmpty($('#m_last_edu').val())){
					alert('최종학력을 선택해주세요');
					$('#m_last_edu').focus();
					return false;
				}

				if(isEmpty($('#m_sido').val())){
					alert('주소(현)를 조회해주세요');
					return false;
				}

				if(isEmpty($('#m_addr').val())){
					alert('주소(현) 번지를 입력해주세요');
					$('#m_addr').focus();
					return false;
				}

				if(!confirm('등록 하시겠습니까?')) return false;
				else return true;
			}

			/*function file_browse()	{
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}*/

			//opener관련 오류가 발생하는 경우 아래 주석을 해지하고, 사용자의 도메인정보를 입력합니다. ("팝업API 호출 소스"도 동일하게 적용시켜야 합니다.)
			//document.domain = "abc.go.kr";
			function jusoCallBack(roadFullAddr,roadAddrPart1,addrDetail,roadAddrPart2,engAddr,jibunAddr,zipNo,admCd,rnMgtSn,bdMgtSn,detBdNmList,bdNm,bdKdcd,siNm,sggNm,emdNm,liNm,rn,udrtYn,buldMnnm,buldSlno,mtYn,lnbrMnnm,lnbrSlno,emdNo,gubun){
				if(gubun === 'juso'){
					$('#m_sido').val(siNm);
					$('#m_gugun').val(sggNm);
					$('#m_dong').val(rn + ' ' + buldMnnm);
					$('#m_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#m_zipcode').val(zipNo);
				}else if(gubun === 'family'){
					$('#m_family_sido').val(siNm);
					$('#m_family_gugun').val(sggNm);
					$('#m_family_dong').val(rn + ' ' + buldMnnm);
					$('#m_family_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#m_family_zip').val(zipNo);
				}
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>" method="post" name="frm" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="9%" >
							<col width="9%" >
							<col width="7%" >
							<col width="9%" >
							<col width="9%" >
							<col width="6%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="6%" >
                            <col width="9%" >
                            <col width="*" >
						</colgroup>
						<tbody>
							<tr>
                                <th>성명(한글)<span style="color:red;">*</span></th>
                                <td colspan="2" class="left">
									<input type='text' name="m_name" id="m_name" style="width:80px" maxlength="20"/>
								</td>
								<th>성명(영문)<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="m_ename" id="m_ename" style="width:80px" maxlength="20"/>
								</td>
                                <th>생년월일<span style="color:red;">*</span></th>
                                <td colspan="2" class="left">
									<input type="text" name="m_birthday" size="10" id="datepicker" style="width:70px;" readonly="true"/>

									<input type="radio" name="m_birthday_id" id="m_birthday_id" value="양" checked/>양
              						<input type="radio" name="m_birthday_id" id="m_birthday_id" value="음"/>음
                                </td>

								<th>주민번호<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<div>
									<input type="text" name="m_person1" id="m_person1" size="6" maxlength="6"  oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									―
									<input type="text" name="m_person2" id="m_person2" size="7" maxlength="7" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									</div>
									<!--<br/>
									<div>성별
									<select name="m_sex" id="m_sex" style="width:50px">
			            				<option value="" selected>선택</option>
										<option value='1'>남</option>
										<option value='2'>여</option>
									</select>
									</div>-->
								</td>
                            </tr>
                            <tr>
              					<th>전화번호</th>
								<td colspan="2" class="left">
									<input type="text" name="m_tel_ddd" id="m_tel_ddd" size="3" maxlength="3" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									-
									<input type="text" name="m_tel_no1" id="m_tel_no1" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									-
									<input type="text" name="m_tel_no2" id="m_tel_no2" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								<th>휴대폰<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="m_hp_ddd" id="m_hp_ddd" size="3" maxlength="3" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									-
									<input type="text" name="m_hp_no1" id="m_hp_no1" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									-
									<input type="text" name="m_hp_no2" id="m_hp_no2" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								<th>비상연락<span style="color:red;">*</span></th>
								<td colspan="2" class="left">
									<input type="text" name="m_emergency_tel" id="m_emergency_tel" size="11" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>&nbsp;("-" 제외)
								</td>
								<th>최종학력<span style="color:red;">*</span></th>
                                <td colspan="2" class="left">
									<select name="m_last_edu" id="m_last_edu" style="width:100px">
										<option value="">선택</option>
										<option value='고등학교'>고등학교</option>
										<option value='전문대'>전문대</option>
										<option value='대학교'>대학교</option>
										<option value='대학원수료'>대학원수료</option>
										<option value='대학원'>대학원</option>
									</select>
                                </td>
							</tr>
                            <tr>
								<th>주소(현)<span style="color:red;">*</span></th>
								<td colspan="8" class="left">
									<input type="text" name="m_zipcode" style="width:40px;" id="m_zipcode" readonly="true"/>
									&nbsp;-&nbsp;
									<input type="text" name="m_sido" id="m_sido" style="width:70px;" readonly="true"/>
									<input type="text" name="m_gugun" id="m_gugun" style="width:80px;" readonly="true"/>
									<input type="text" name="m_dong" id="m_dong" style="width:100px;" readonly="true"/>
									<input type="text" name="m_addr" id="m_addr" style="width:330px;" notnull errname="번지" onKeyUp="checklength(this,50)" />
									<span>
										<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=juso','family_zip_select','scrollbars=yes,width=600,height=400')">조회</a>
									<span>
                                </td>
								<th>결혼기념일</th>
                                <td colspan="2" class="left">
									<input name="m_marry_date" type="text" size="10" id="datepicker1" style="width:70px;" readonly="true"/>
								</td>
                            </tr>
                         	<tr>
                                <th class="first">경조가입여부</th>
                                <td colspan="2" class="left">
									<input type="radio" name="m_sawo_id" id="m_sawo_id" value="Y" checked/>가입
              						<input type="radio" name="m_sawo_id" id="m_sawo_id" value="N"/>안함
								&nbsp;
								</td>
								<th>취미</th>
                                <td colspan="2" class="left">
									<input name="m_hobby" id="m_hobby" type="text" id="emp_hobby" size="13"/>
								</td>
								<th>종교</th>
								<td colspan="2" class="left">
									<input name="m_faith" id="m_faith" type="text" id="emp_faith" style="width:50px"/>
								</td>
								<th>장애/등급</th>
								<td colspan="2" class="left">
                				<%
                				objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '22' ORDER BY emp_etc_code ASC "

								Set rsEtc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="m_disabled" id="m_disabled" style="width:90px">
                  						<option value="">선택</option>
                				<%
								Do Until rsEtc.EOF
			  				  	%>
                					<option value='<%=rsEtc("emp_etc_name")%>'><%=rsEtc("emp_etc_name")%></option>
                				<%
									rsEtc.MoveNext()
								Loop
								rsEtc.Close() : Set rsEtc = Nothing
							  	%>
            						</select>
									-
									<select name="m_disab_grade" id="m_disab_grade" style="width:50px">
										<option value="">선택</option>
										<option value='1급'>1급</option>
										<option value='2급'>2급</option>
										<option value='3급'>3급</option>
										<option value='4급'>4급</option>
										<option value='5급'>5급</option>
										<option value='6급'>6급</option>
										<option value='중증'>중증</option>
										<option value='경증'>경증</option>
                					</select>
								</td>
                 			</tr>
                            <tr>
								<th>병역유형</th>
								<td colspan="2" class="left">
                				<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '06' ORDER BY emp_etc_code ASC "

								Set rsMiliId = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="m_military_id" id="m_military_id" style="width:90px">
                  						<option value="" selected>선택</option>
                				<%
								Do Until rsMiliId.EOF
			  				  	%>
                						<option value='<%=rsMiliId("emp_etc_name")%>'><%=rsMiliId("emp_etc_name")%></option>
                				<%
									rsMiliId.MoveNext()
								Loop
								rsMiliId.Close() : Set rsMiliId = Nothing
							  	%>
									</select>
								</td>
								<th>병역계급</th>
								<td colspan="2" class="left">
                				<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '07' ORDER BY emp_etc_code ASC "

								Set rsMiliGrade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
									<select name="m_military_grade" id="m_military_grade" style="width:90px">
                  						<option value="" selected>선택</option>
                				<%
								Do Until rsMiliGrade.EOF
			  				  	%>
                						<option value='<%=rsMiliGrade("emp_etc_name")%>'><%=rsMiliGrade("emp_etc_name")%></option>
                				<%
									rsMiliGrade.MoveNext()
								Loop
								rsMiliGrade.Close() : Set rsMiliGrade = Nothing
								DBConn.Close() : Set DBConn = Nothing
							  	%>
                					</select>
								</td>
								<th>복무기간</th>
								<td colspan="2" class="left">
									<input name="m_military_date1" type="text" size="10" id="datepicker2" style="width:70px;" readonly="true"/>
									∼
									<input name="m_military_date2" type="text" size="10" id="datepicker3" style="width:70px;" readonly="true"/>
								</td>
								<th>면제사유</th>
								<td colspan="2" class="left">
									<input name="m_military_comm" type="text" id="m_military_comm" size="13"/>
								</td>
							</tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th scope="row">사진등록</th>
								<td class="left">
									<input type="file" name= "att_file" size="70" accept="image/gif" /> * 첨부파일은 1개만 가능하며 최대용량은 2MB
                                </td>
							</tr>
						</tbody>
                    </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div class="btnCenter">
                         <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" /></span>
                         <span class="btnType01"><input type="button" value="취소" onclick="goBefore();" /></span>
                    </div>
                    </td>
				    <td width="52%">
					<div class="btnCenter">
                    <span class="btnType04" style="width:710px;">기본 정보 등록 후 ☞ 가족사항 ☞ 학력사항 ☞ 경력사항 ☞ 자격사항 ☞ 교육사항 ☞ 어학능력을 등록하시기 바랍니다.</span>
					</div>
                    </td>
			      </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>