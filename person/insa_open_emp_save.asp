<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->

<!--include file="xmlrpc.asp"-->
<!--include file="class.EmmaSMS.asp"-->
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
'### Upload Form
'===================================================
Dim UploadForm
Set UploadForm = Server.CreateObject("ABCUpload4.XForm")

UploadForm.AbsolutePath = True
UploadForm.Overwrite = True
UploadForm.MaxUploadSize = 1024*1024*50

'===================================================
'### Request & Params
'===================================================
Dim emp_name, emp_ename, emp_birthday, emp_birthday_id
Dim emp_tel_ddd, emp_tel_no1, emp_tel_no2
Dim emp_hp_ddd, emp_hp_no1, emp_hp_no2
Dim emp_family_zip, emp_family_sido, emp_family_gugun, emp_family_dong, emp_family_addr
Dim emp_emergency_tel, emp_zipcode, emp_sido, emp_gugun, emp_dong, emp_addr, emp_email
Dim emp_marry_date, emp_hobby, emp_military_id, emp_military_grade, emp_military_date1, emp_military_date2
Dim emp_military_comm, emp_faith, emp_extension_no, emp_last_edu, kone_email, emp_hp
Dim path, filenm, filename, fileType, save_path, curr_date, err_state, end_msg
Dim rsPerson, arrPerson, be_pg

emp_no = UploadForm("emp_no")
emp_name = UploadForm("emp_name")
emp_ename = UploadForm("emp_ename")

emp_birthday = UploadForm("emp_birthday")
If f_toString(emp_birthday, "") = "" Then
	emp_birthday = "1900-01-01"
End If

emp_birthday_id = UploadForm("emp_birthday_id")
emp_tel_ddd = UploadForm("emp_tel_ddd")
emp_tel_no1 = UploadForm("emp_tel_no1")
emp_tel_no2 = UploadForm("emp_tel_no2")
emp_hp_ddd = UploadForm("emp_hp_ddd")
emp_hp_no1 = UploadForm("emp_hp_no1")
emp_hp_no2 = UploadForm("emp_hp_no2")
emp_family_zip = UploadForm("emp_family_zip")
emp_family_sido = UploadForm("emp_family_sido")
emp_family_gugun = UploadForm("emp_family_gugun")
emp_family_dong = UploadForm("emp_family_dong")
emp_family_addr = UploadForm("emp_family_addr")
emp_emergency_tel = UploadForm("emp_emergency_tel")
emp_zipcode = UploadForm("emp_zipcode")
emp_sido = UploadForm("emp_sido")
emp_gugun = UploadForm("emp_gugun")
emp_dong = UploadForm("emp_dong")
emp_addr = UploadForm("emp_addr")
emp_email = UploadForm("emp_email")

emp_marry_date = UploadForm("emp_marry_date")
If f_toString(emp_marry_date, "") = "" Then
	emp_marry_date = "1900-01-01"
End If

emp_hobby = UploadForm("emp_hobby")

emp_military_id = UploadForm("emp_military_id")
emp_military_grade = UploadForm("emp_military_grade")

emp_military_date1 = UploadForm("emp_military_date1")
If f_toString(emp_military_date1, "") = "" Then
	emp_military_date1 = "1900-01-01"
End If

emp_military_date2 = UploadForm("emp_military_date2")
If f_toString(emp_military_date2, "") = "" Then
	emp_military_date2 = "1900-01-01"
End If

emp_military_comm = UploadForm("emp_military_comm")
emp_faith = UploadForm("emp_faith")
emp_extension_no = UploadForm("emp_extension_no")
emp_last_edu = UploadForm("emp_last_edu")
mg_group = UploadForm("mg_group")

be_pg = "/person/insa_individual_emp_add.asp"

path = Server.MapPath ("/emp_photo")

Set filenm = UploadForm("att_file")(1)
filename = filenm

If filenm <> "" Then
	filename = filenm.safeFileName
	fileType = Mid(filename, InStrRev(filename, ".") + 1)
	filename = emp_name & "_" & emp_no & "_" & "photo." & fileType
	save_path = path & "\" & filename
End If

If filenm.length > 1024*1024*8  Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('파일 용량은 2MB를 넘을 수 없습니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

' 로그인 정보(memb) 변경
kone_email = emp_email & "@k-one.co.kr"
emp_hp = emp_hp_ddd & "-" & emp_hp_no1 & "-" & emp_hp_no2

'curr_date = mid(cstr(now()),1,10)
curr_date = f_FormatDate()

If filenm <> "" Then
   filenm.save save_path
End If

objBuilder.Append "CALL USP_PERSON_INDIVIDUAL_UPDATE('"&emp_no&"', '"&emp_ename&"', '"&emp_birthday&"', "
objBuilder.Append "	'"&emp_birthday_id&"', '"&emp_family_zip&"', '"&emp_family_sido&"', "
objBuilder.Append "	'"&emp_family_gugun&"', '"&emp_family_dong&"', '"&emp_family_addr&"', "
objBuilder.Append " '"&emp_zipcode&"', '"&emp_sido&"', '"&emp_gugun&"', "
objBuilder.Append " '"&emp_dong&"', '"&emp_addr&"', '"&emp_tel_ddd&"', "
objBuilder.Append "	'"&emp_tel_no1&"', '"&emp_tel_no2&"', '"&emp_hp_ddd&"', "
objBuilder.Append "	'"&emp_hp_no1&"', '"&emp_hp_no2&"', '"&emp_email&"', "
objBuilder.Append "	'"&emp_military_id&"', '"&emp_military_date1&"', '"&emp_military_date2&"', "
objBuilder.Append "	'"&emp_military_grade&"', '"&emp_military_comm&"', '"&emp_hobby&"', "
objBuilder.Append "	'"&emp_faith&"', '"&emp_last_edu&"', '"&emp_marry_date&"', "
objBuilder.Append "	'"&emp_emergency_tel&"', '"&emp_extension_no&"', '"&filename&"', "
objBuilder.Append "	'"&emp_name&"', '"&emp_hp&"', '"&kone_email&"', '"&curr_date&"');"

Call Rs_Open(rsPerson, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsPerson.EOF Then
	arrPerson = rsPerson.getRows()

	err_state = arrPerson(0, 0)
End If

Call Rs_Close(rsPerson)
DBConn.Close() : Set DBConn = Nothing

If Err.number <> 0 Or err_state <> "0" Then
	end_msg = "System Error가 발생하였습니다.\n담당자 혹은 관리자에게 문의해 주세요."
Else
	end_msg = "정상적으로 등록되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&be_pg&"');"
Response.Write "</script>"
Response.End
%>

