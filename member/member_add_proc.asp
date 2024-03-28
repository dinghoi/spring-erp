<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
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
Dim m_ename, m_birthday, m_birthday_id, m_person1, m_person2
Dim m_sex, m_tel_ddd, m_tel_no1, m_tel_no2, m_hp_ddd, m_hp_no1, m_hp_no2
Dim m_emergency_tel, m_last_edu, m_sido, m_gugun, m_dong, m_addr, m_zipcode
Dim m_marry_date, m_sawo_id, m_hobby, m_faith, m_disabled, m_disab_grade
Dim m_military_id, m_military_grade, m_military_date1, m_military_date2
Dim m_military_comm, att_file, sex_id, m_disabled_yn, path
Dim curr_date, filename, filenm, fileType, save_path
Dim uploadForm, rsMem, end_msg, rsSeq, fileIdx

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")

uploadForm.AbsolutePath = True
uploadForm.Overwrite = True
uploadForm.MaxUploadSize = 1024*1024*50

m_name = uploadForm("m_name")
m_ename = uploadForm("m_ename")
m_birthday = uploadForm("m_birthday")
m_birthday_id = uploadForm("m_birthday_id")
m_person1 = uploadForm("m_person1")
m_person2 = uploadForm("m_person2")
m_sex = uploadForm("m_sex")
m_tel_ddd = uploadForm("m_tel_ddd")
m_tel_no1 = uploadForm("m_tel_no1")
m_tel_no2 = uploadForm("m_tel_no2")
m_hp_ddd = uploadForm("m_hp_ddd")
m_hp_no1 = uploadForm("m_hp_no1")
m_hp_no2 = uploadForm("m_hp_no2")
m_emergency_tel = uploadForm("m_emergency_tel")
m_last_edu = uploadForm("m_last_edu")
m_sido = uploadForm("m_sido")
m_gugun = uploadForm("m_gugun")
m_dong = uploadForm("m_dong")
m_addr = uploadForm("m_addr")
m_zipcode = uploadForm("m_zipcode")
m_marry_date = uploadForm("m_marry_date")
m_sawo_id = uploadForm("m_sawo_id")
m_hobby = uploadForm("m_hobby")
m_faith = uploadForm("m_faith")
m_disabled = uploadForm("m_disabled")
m_disab_grade = uploadForm("m_disab_grade")
m_military_id = uploadForm("m_military_id")
m_military_grade = uploadForm("m_military_grade")
m_military_date1 = uploadForm("m_military_date1")
m_military_date2 = uploadForm("m_military_date2")
m_military_comm = uploadForm("m_military_comm")
att_file = uploadForm("att_file")

If m_person2 <> "" Then
	sex_id = Mid(CStr(m_person2), 1, 1)

	If (sex_id = "1" Or sex_id = "3" Or sex_id = "5" Or sex_id = "7" Or sex_id ="9") Then
		m_sex = "남"
	Else
		m_sex = "여"
	End If
End If

If m_disabled = "해당사항없음" Or m_disabled = "" Then
	m_disabled_yn = "N"
	m_disab_grade = ""
Else
	m_disabled_yn = "Y"
End If

'파일 업로드 설정
path = Server.MapPath("/emp_photo")

Set filenm = uploadForm("att_file")(1)
filename = filenm

If filenm.FileExists Then
	If filenm.length > 1024*1024*8  Then
		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('업로드 파일 최대 용량은 2MB를 넘을 수 없습니다.');"
		Response.Write "	history.go(-1);"
		Response.Write "</script>"
		Response.End
	Else
		objBuilder.Append "SELECT m_seq FROM member_info ORDER BY m_seq DESC LIMIT 1;"

		Set rsSeq = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsSeq.EOF Or rsSeq.BOF Then
			fileIdx = "1"
		Else
			fileIdx = rsSeq("m_seq") + 1
		End If
		rsSeq.Close() : Set rsSeq = Nothing

		filename = filenm.safeFileName
		fileType = Mid(filename, InStrRev(filename, ".") + 1)
		filename = m_name&"_"&Replace(Mid(CStr(Now()), 1, 10), "-", "")&"_"&fileIdx&"."&fileType
		save_path = path&"\"&filename

		filenm.save save_path
	End If
End If

If f_toString(m_birthday, "") = "" Then
	m_birthday = "1900-01-01"
End If

If f_toString(m_military_date1, "") = "" Then
	m_military_date1 = "1900-01-01"
End If

If f_toString(m_military_date2, "") = "" Then
	m_military_date2 = "1900-01-01"
End If

If f_toString(m_marry_date, "") = "" Then
	m_marry_date = "1900-01-01"
End If

DBConn.BeginTrans

objBuilder.Append "INSERT INTO member_info(m_name, m_ename, m_birthday, m_birthday_id, m_person1, m_person2,"
objBuilder.Append "m_sex, m_tel_ddd, m_tel_no1, m_tel_no2, m_hp_ddd, m_hp_no1, m_hp_no2, "
objBuilder.Append "m_zipcode, m_sido, m_gugun, m_dong, m_addr, m_emergency_tel, m_sawo_id, "
objBuilder.Append "m_hobby, m_disabled, m_disab_grade, m_military_id, m_military_grade, "
objBuilder.Append "m_military_date1, m_military_date2, m_military_comm, m_marry_date, "
objBuilder.Append "m_faith, m_last_edu, m_image)"
objBuilder.Append "VALUES('"&m_name&"', '"&m_ename&"', '"&m_birthday&"', '"&m_birthday_id&"', '"&m_person1&"', '"&m_person2&"', "
objBuilder.Append "'"&m_sex&"', '"&m_tel_ddd&"', '"&m_tel_no1&"', '"&m_tel_no2&"', '"&m_hp_ddd&"', '"&m_hp_no1&"', '"&m_hp_no2&"', "
objBuilder.Append "'"&m_zipcode&"', '"&m_sido&"', '"&m_gugun&"', '"&m_dong&"', '"&m_addr&"', '"&m_emergency_tel&"', '"&m_sawo_id&"', "
objBuilder.Append "'"&m_hobby&"', '"&m_disabled&"', '"&m_disab_grade&"', '"&m_military_id&"', '"&m_military_grade&"', "
objBuilder.Append "'"&m_military_date1&"', '"&m_military_date2&"', '"&m_military_comm&"', '"&m_marry_date&"', "
objBuilder.Append "'"&m_faith&"', '"&m_last_edu&"', '"&filename&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생했습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 등록되었습니다.\n기타 사항(가족사항 등)을 추가로 입력해주세요."

	'입력 상태 유지 시퀀스 번호 설정
	objBuilder.Append "SELECT m_seq, m_name FROM member_info ORDER BY m_reg_date DESC LIMIT 1;"

	Set rsMem = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Response.cookies("nkp_member")("coo_m_seq") = rsMem("m_seq")
	Response.cookies("nkp_member")("coo_m_name") = rsMem("m_name")

	rsMem.Close() : Set rsMem = Nothing
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.href='/member/member_family.asp';"
Response.Write"</script>"
Response.End
%>