<%
'===========================================================================
'  Class ����
'===========================================================================

'---------------------------------------------------------------------------
'Ŭ���� �� : StringBuilder
' Description : StringBuilder Ŭ����ȭ
' Example :
    'Dim str
    'Set str = New StringBuilder

    'str.Append "�ȳ��ϼ���."
    'str.Append "�ݰ����ϴ�."
    'Response.Write str.ToString()

    'str.Clear
    'str.Append "�ι�°�Դϴ�."
    'str.Append "����."
    'Response.Write str.ToString()

    'str.Clear()
'---------------------------------------------------------------------------
Class StringBuilder

    Private mArr        '�����ų ���ڿ��� ���� �迭
    Private mGrowthRate '�迭 ũ�� ���� ��ų ����
    Private mItemCount  '�迭 ��� ����
    Public  JoinString  '��ĥ�� ����� ����

    Private Sub Class_Initialize()
        JoinString= ""
        Call Clear()
    End Sub

    Private Sub Class_Terminate()
        'Erase Buffer
    End Sub

    Public Sub Append(ByVal strValue)
        If mItemCount > UBound(mArr) Then
            ReDim Preserve mArr(UBound(mArr) + mGrowthRate)
        End If
        mArr(mItemCount) = strValue
        mItemCount = mItemCount + 1
    End Sub

    Public Function ToString()
        If mItemCount > 0 Then
            ToString = Join(mArr, JoinString)
        Else
            ToString = ""
        End If
    End Function

    Public Sub Clear()
        mGrowthRate = 50
        mItemCount = 0
        ReDim mArr(mGrowthRate)
    End Sub

End Class

'===========================================================================
'   Sub �Լ� ����
'===========================================================================

'---------------------------------------------------------------------------
' Function Name : SetDomainCookie
' Description : Cookie �� ����
'---------------------------------------------------------------------------
Sub SetDomainCookie(ByVal key, ByVal val, ByVal domain)
    Response.Cookies(key) = val

    If domain <> "" Then
        Response.Cookies(key).Domain = domain
        Response.Cookies(key).Path = "/"
    End If
End Sub

'---------------------------------------------------------------------------
' Function Name : MakeTempPwd
' Description : �ӽú�й�ȣ
'---------------------------------------------------------------------------
Sub MakeTempPwd(ByRef pwd)
    Randomize Timer
    Const StringTable = ("0123456789abcdefghijklmnopqrstuvwxyz")

    For i = 1 To 7
        pwd = pwd & Mid(StringTable, Int(Rnd(1)*Len(StringTable))+1,1)
    Next
End Sub

'---------------------------------------------------------------------------
' * Check Function
'---------------------------------------------------------------------------
Sub Formchk()
    Dim key

    For Each key in Request.Form
        Response.Write key & " = " & "Trim(Request.Form("""& key &""")) " & "<BR>"
    Next

    Response.WRite "<HR>"

    For Each key in Request.Form
        Response.Write key & " = " & Request.Form(key) & "<BR>"
    Next
End Sub

Sub Querychk()
    Dim key

    For Each key in Request.QueryString
        Response.Write key & " = " & "Trim(Request.QueryString("""& key &"""))" & "<BR>"
    Next

    Response.WRite "<HR>"

    For Each key in Request.QueryString
        Response.Write key & " = " & Request.QueryString(key) & "<BR>"
    Next
End Sub

Sub Cookiechk()
    Dim key, dickey

    For Each key in Request.Cookies
        If Request.Cookies(key).HasKeys Then
            '��ųʸ��� �ִ� ��� Ű���� �˻��ϱ� ���� �� �ٸ� For Each �� ���
            For Each dickey in Request.Cookies(key)
                Response.Write "(" & key & ")(" & dickey & ") = " & Request.Cookies(key)(dickey) & "<BR>"
            Next
        Else
            '�Ϲ� ��Ű
            Response.Write key & " = " & Request.Cookies(key) & "<BR>"
        End If
    Next

    Response.Write "<HR>"

    For Each key In Request.Cookies
        If Request.Cookies(key).HasKeys Then
            '��ųʸ��� �ִ� ��� Ű���� �˻��ϱ� ���� �� �ٸ� For Each �� ���
            For Each dickey In Request.Cookies(key)
                Response.Write  dickey & " = Request.Cookies("""& key &""")("""& dickey &""") " & "<BR>"
            Next
        Else
            '�Ϲ� ��Ű
            Response.Write key & " = Request.Cookies("""& key &""")" & "<BR>"
        End If
    Next
End Sub

'----------------- ���̺� Į���� + ���ڵ�� �� ���� ------------------------------
Sub RS_Column(tablename, strConnect)
    Dim oRs, fldTable

    Set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.Open tablename, strConnect, 1

    If oRs.State = 1 Then
        If NOT oRs.EOF Then
            For Each fldTable In oRs.Fields
                Response.WRite fldTable.name & " = oRs(""" & fldTable.name & """) " & "<BR>"
            Next

            Response.Write "<HR>"

            For Each fldTable In oRs.Fields
                Response.WRite fldTable.name & " = " & fldTable.value & "<BR>"
            Next

            Response.Write "<HR>"

            For Each fldTable In oRs.Fields
                Response.WRite fldTable.name & " = " & "RS(""" & fldTable.name & """) " & "<BR>"
            Next
        End If

        oRs.Close
        Set oRs = Nothing
    End If
End Sub

'---------------------------------------------------------------------------
' Function Name : SelectBoxEmpOrg
' Description : ������ �˻�
' Author : ����ȣ
' DATE : 2021-05-21
' param :
'   SelectEmpOrgList(name, id, css, ������)
'---------------------------------------------------------------------------
Sub SelectEmpOrgList(name, id, css, condi)
	Dim SQL, rsOrg

	'SQL = "SELECT org_name "
	'SQL = SQL & "FROM emp_org_mst "
	'SQL = SQL & "WHERE (ISNULL(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00') "
	'SQL = SQL & "	AND org_level = 'ȸ��' "
	'SQL = SQL & "ORDER BY FIELD(org_company, '���̿�') DESC, org_code DESC "
    SQL = "CALL USP_COMM_ORG_MST_INFO()"

	'Set rsOrg = Server.CreateObject("ADODB.RecordSet")
	'rsOrg.Open SQL, DBConn, 1
    Set rsOrg = DBConn.Execute(SQL)

	Response.Write "<select name='"&name&"' id='"&id&"' style='"&css&"'>"

    'If SysAdminYn = "Y" Then
        Response.Write "<option value='��ü' "
        If condi = "��ü" Then
            Response.Write "selected"
        End If
        Response.Write ">��ü</option>"
    'End If

	Do Until rsOrg.EOF
		Response.Write "<option value='"&rsOrg("org_name")&"' "

		If condi = rsOrg("org_name") Then
			Response.Write "selected"
		end If
		Response.Write ">"&rsOrg("org_name")&"</option> "

		rsOrg.MoveNext
	Loop
	'rsOrg.Close() : Set rsOrg = Nothing
    Call Rs_Close(rsOrg)
	Response.Write "</select>"
End Sub

'---------------------------------------------------------------------------
' Function Name : SelectEmpOrgLevel
' Description : ���� ������ SelectBox
' Author : ����ȣ
' DATE : 2021-08-09
' Param : name ��, id ��, stylesheet, ����
'---------------------------------------------------------------------------
Sub SelectEmpOrgLevel(name, id, css, condi)
	Dim SQL, rsOrg

    SQL = "CALL USP_COMM_ORG_LEVEL_INFO"
    Set rsOrg = DBConn.Execute(SQL)

	Response.Write "<select name='"&name&"' id='"&id&"' style='"&css&"'>"

	Do Until rsOrg.EOF
		Response.Write "<option value='"&rsOrg("org_name")&"' "

		If condi = rsOrg("org_name") Then
			Response.Write "selected"
		end If
		Response.Write ">"&rsOrg("org_name")&"</option> "

		rsOrg.MoveNext
	Loop
	rsOrg.Close() : Set rsOrg = Nothing

	Response.Write "</select>"
End Sub

'RecordSet ��ü Open
Sub Rs_Open(rs_name, db_name, query)
    Set rs_name = Server.CreateObject("ADODB.RecordSet")
    rs_name.Open query, db_name, 1
End Sub

'RecordSet ��ü Close
Sub Rs_Close(rs_name)
    rs_name.Close() : Set rs_name = Nothing
End Sub

'Command ��ü Open
Sub Cmd_Open(str)
    Set str = Server.CreateObject("ADODB.Command")
End Sub

'Command ��ü Close
Sub Cmd_Close(str)
    Set str = Nothing
End Sub

'---------------------------------------------------------------------------
' Function Name : EmpOrgText
' Description : ������ Text(����� ����)
' Author : ����ȣ
' DATE : 2021-05-24
'---------------------------------------------------------------------------
Sub EmpOrgText(company, bonbu, team)
	Response.Write Company

	If bonbu <> "" And Not IsNull(bonbu) Then
		Response.Write " - " & bonbu
	End If

	If team <> "" And Not IsNull(team) Then
		Response.Write " - " & team
	End If
End Sub

'---------------------------------------------------------------------------
' Function Name : EmpOrgInSaupbuText
' Description : ������ Text(����� ����)
' Author : ����ȣ
' DATE : 2021-06-01
'---------------------------------------------------------------------------
Sub EmpOrgInSaupbuText(company, bonbu, saupbu, team)
	Response.Write Company

	If bonbu <> "" And Not IsNull(bonbu) Then
		Response.Write " - " & bonbu
	End If

	If saupbu <> "" And Not IsNull(saupbu) Then
		Response.Write " - " & saupbu
	End If

	If team <> "" And Not IsNull(team) Then
		Response.Write " - " & team
	End If
End Sub

'---------------------------------------------------------------------------
' Function Name : EmpOrgInSaupbuText
' Description : ������ Text(����� ����)
' Author : ����ȣ
' DATE : 2021-06-01
' Param : �����ڵ�
'---------------------------------------------------------------------------
Sub EmpOrgCodeSelect(code)
	Dim SQL, rsOrg, arrOrg
	Dim oCompany, oBonbu, oSaupbu, oTeam

    SQL = "CALL USP_COMM_ORG_SELECT_INFO('"&code&"')"

    Call Rs_Open(rsOrg, DBConn, SQL)

	If Not rsOrg.EOF Then
        arrOrg = rsOrg.getRows()

		'oCompany = rsOrg("org_company")
		'oBonbu = rsOrg("org_bonbu")
		'oSaupbu = rsOrg("org_saupbu")
		'oTeam = rsOrg("org_team")
        oCompany = arrOrg(0, 0)
        oBonbu = arrOrg(1, 0)
		oSaupbu = arrOrg(2, 0)
		oTeam = arrOrg(3, 0)

		Response.Write oCompany

		If oBonbu <> "" And Not IsNull(oBonbu) Then
			Response.Write " > " & oBonbu
		End If

		If oSaupbu <> "" And Not IsNull(oSaupbu) Then
			Response.Write " > " & oSaupbu
		End If

		If oTeam <> "" And Not IsNull(oTeam) Then
			Response.Write " > " & oTeam
		End If
	Else
        Response.Write ""
	End If

    Call Rs_Close(rsOrg)
End Sub

'---------------------------------------------------------------------------
' Function Name : SelectBoxEmpOrg
' Description : ������ �˻�
' Author : ����ȣ
' DATE : 2021-05-21
'---------------------------------------------------------------------------
Sub SelectEmpEtcCodeList(name, id, css, etc_type, condi)
	Dim rsEtcCode, arrEtcCode, emp_etc_name, i

    objBuilder.Append "CALL USP_COMM_ETC_CODE_INFO('"&etc_type&"')"

    Call Rs_Open(rsEtcCode, DBConn, objBuilder.ToString())
    objBuilder.Clear()

    arrEtcCode = rsEtcCode.getRows()

	Response.Write "<select name='"&name&"' id='"&id&"' style='"&css&"'>"
	Response.Write "<option value='' "

	If condi = "" Then
		Response.Write "selected"
	End If

	Response.Write ">����</option>"

    For i = LBound(arrEtcCode) To UBound(arrEtcCode, 2)
        emp_etc_name = arrEtcCode(0, i)

        Response.Write "<option value='"&emp_etc_name&"' "

		If condi = emp_etc_name Then
			Response.Write "selected"
		End If

		Response.Write ">"&emp_etc_name&"</option> "
    Next

	Call Rs_Close(rsEtcCode)

	Response.Write "</select>"
End Sub

'---------------------------------------------------------------------------
' Function Name : ViewExcelType
' Description : ���� DownLoad ���� �ڵ�
' Author : ����ȣ
' DATE : 2021-06-15
'---------------------------------------------------------------------------

Sub ViewExcelType(filename)
	Response.Buffer = True
	Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment; filename="&filename
End Sub

'---------------------------------------------------------------------------
' Function Name : Page_Navi
' Description : page navigator
' Author : ����ȣ
' DATE : 2021-07-20
'---------------------------------------------------------------------------

Sub Page_Navi(page, url, param, tot_page)
    Dim intstart, intend, first_page, i

	intstart = (Int((page - 1) / 10) * 10) + 1
	intend = intstart + 9
	first_page = 1

	If intend > tot_page Then
		intend = tot_page
	End If

	Response.Write "<div id='paging'>"
	Response.Write "<a href="&be_pg&"?page="&first_page&param&">[ó��]</a>"

	If intstart > 1 Then
		Response.Write "<a href="&be_pg&"?page="&intstart - 1&param&">[����]</a>"
	End If

	For i = intstart To intend
		If i = Int(page) Then
			Response.Write "<b>["&i&"]</b>"
		Else
			Response.Write "<a href="&be_pg&"?page="&i&param&">["&i&"]</a>"
		End If
	Next
	If intend < tot_page Then
		Response.Write "<a href="&be_pg&"?page="&intend + 1&param&">[����]</a> "
		Response.Write "<a href="&be_pg&"?page="&tot_page&param&">[������]</a>"
	Else
		Response.Write "[����]&nbsp;[������]"
	End If

	Response.Write "</div>"
End Sub

'---------------------------------------------------------------------------
' Function Name : Page_Navi
' Description : page navigator
' Author : ����ȣ
' DATE : 2021-07-20
'---------------------------------------------------------------------------

Sub Page_Navi_Ver2(page, url, param, tot_record, pgsize)
    Dim intstart, intend, first_page, i
    Dim total_page

    'Result.PageCount
    If tot_record Mod pgsize = 0 Then
	    total_page = Int(tot_record / pgsize)
    Else
        total_page = Int((tot_record / pgsize) + 1)
    End If

	intstart = (Int((page - 1) / 10) * 10) + 1
	intend = intstart + 9
	first_page = 1

	If intend > total_page Then
		intend = total_page
	End If

	Response.Write "<div id='paging'>"
	Response.Write "<a href="&be_pg&"?page="&first_page&param&">[ó��]</a>"

	If intstart > 1 Then
		Response.Write "<a href="&be_pg&"?page="&intstart - 1&param&">[����]</a>"
	End If

	For i = intstart To intend
		If i = Int(page) Then
			Response.Write "<b>["&i&"]</b>"
		Else
			Response.Write "<a href="&be_pg&"?page="&i&param&">["&i&"]</a>"
		End If
	Next

	If intend < total_page Then
		Response.Write "<a href="&be_pg&"?page="&intend + 1&param&">[����]</a> "
		Response.Write "<a href="&be_pg&"?page="&total_page&param&">[������]</a>"
	Else
		Response.Write "[����]&nbsp;[������]"
	End If

	Response.Write "</div>"
End Sub

'---------------------------------------------------------------------------
' Function Name : EmpInfoName
' Description : ������ Select
' Author : ����ȣ
' DATE : 2021-08-20
'---------------------------------------------------------------------------
Sub EmpInfo_Name(id)
	Dim rsInfo, arrInfo

    objBuilder.Append "CALL USP_COMM_EMP_MASTER_NAME('"&id&"')"
    Set rsInfo = DBConn.Execute(objBuilder.ToString())
    objBuilder.Clear()

    arrInfo = rsInfo.getRows()

    Call Rs_Close(rsInfo)

    Response.Write arrInfo(0, 0)
End Sub

'====================================
' Function Name : CostEndError
' Description : ��븶�� ���� �� �̵�
' Author : ����ȣ
' DATE : 2021-10-07
'====================================
Sub CostEndError(url)
    Response.Write "<script type='text/javascript'>"
    Response.Write "	alert('ó���� Error�� �߻��Ͽ����ϴ�.');"
    Response.Write "	location.replace('"&url&"');"
    Response.Write "</script>"
    Response.End
End Sub

'====================================
' Function Name : f_SalesCompany
' Description : ���� ȸ�� ����
' Author : ����ȣ
' DATE : 2022-02-21
'====================================
Function f_SalesCompany(name)
    Dim company

    Select Case name
        Case "���̿��������", "(��)���̿��������", "�ֽ�ȸ�� ���̿��������", "(��)���̿�", "�ֽ�ȸ�� ���̿�"
            company = "���̿�"
        Case "(��)���̳�Ʈ����", "�ֽ�ȸ�� ���̳�Ʈ����"
            company = "���̳�Ʈ����"
        Case "(��)���̽ý���", "�ֽ�ȸ�� ���̽ý���"
            company = "���̽ý���"
        Case Else
            company = name
    End Select

    f_SalesCompany = company
End Function

'===========================================================================
'   Function ����
'===========================================================================
'---------------------------------------------------------------------------
'�Լ��� : f_EncUft8
'INPUT : ���ڿ�
'��ɼ��� : UTF-8 Encode
'Example :
'---------------------------------------------------------------------------
Function f_EncUft8(astr)
    Dim utftext, n, c

    utftext = ""

    For n = 1 To Len(astr)
        c = AscW(Mid(astr, n, 1))

        If c < 128 Then
            utftext = utftext + Mid(astr, n, 1)
        ElseIf c > 127 And c < 2048 Then
            utftext = utftext + Chr((c \ 64) Or 192)
            utftext = utftext + Chr((c And 63) Or 128)
        Else
            utftext = utftext + Chr((c \ 144) Or 234)
            utftext = utftext + Chr(((c \ 64) And 63) Or 128)
            utftext = utftext + Chr((c And 63) Or 128)
        End If
    Next
    f_EncUft8 = utftext
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_Request
'INPUT : (Request ���� ����)
'��ɼ��� : Reqeuest GET or FORM
'Example :
'---------------------------------------------------------------------------

Function f_Request(str)
	Dim rVal

	rVal = Request.QueryString(str)

	If f_toString(rVal, "0") = "0" Then
		rVal = Request.Form(str)
	End If

	f_Request = rVal

End Function

'---------------------------------------------------------------------------
'�Լ��� : f_toString
'INPUT : (����,���ڰ� NULL Ȥ�� �����ϰ�� ��ü�� ��)
'��ɼ��� : Null Ȥ�� ���� üũ
'Example :
'---------------------------------------------------------------------------
Function f_toString(str1, str2)
	If IsNull(str1) Or str1 = "" Or IsEmpty(str1) Then
		str1 = ""
		If Not IsNull(str2) And str2 <> "" And Not IsEmpty(str2) Then
			str1 = str2
		End If
	End If

	f_toString = str1
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_getRsToDic
'INPUT :
'��ɼ��� : recordset to dictionary
'Example :
'---------------------------------------------------------------------------
'Public Function func_GetRsToDic(rs)
Function f_getRsToDic(rs)
	Dim rsRow				'��� ���ڵ�� 2���� �迭
	Dim rsCount				'���ڵ�� rows
	Dim resultList			'������ ����Ʈ
	Dim dic					'���ڵ� row dictionary
	Dim i, j

	Set resultList = Server.CreateObject("Scripting.Dictionary")

	If Not rs.EOF Then
		rsRow = rs.GetRows()
		rsCount = Ubound(rsRow,2)

		For i = 0 To rsCount
			Set dic = Server.CreateObject("Scripting.Dictionary")

			For j = 0 To rs.fields.count - 1
				name = rs.fields(j).name

				If isnull(rsRow(j, i)) Then
					value = ""
				Else
					value = CStr(rsRow(j, i))
				End If

				dic.add name, value
			Next

			resultList.add i, dic
            dic.RemoveAll()
		Next
	End if

	Set f_getRsToDic = resultList
    resultList.RemoveAll()
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_stripTags
'INPUT : htmlDoc
'��ɼ��� : HTML �±�����
'---------------------------------------------------------------------------
Function f_stripTags(htmlDoc)
    Dim rex
    Set rex = New Regexp

    rex.Pattern = "<[^>]+>"
    rex.Global = True
    f_stripTags = rex.Replace(htmlDoc,"")
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_DB_IN_STR
'INPUT : cur_str ==> �˻��� ���ڿ�
'��ɼ��� : DB�Է��Ҷ� ' �� '' �� ��ü
'---------------------------------------------------------------------------
Function f_DB_IN_STR(cur_str)
    If Not isNull(cur_str) Then
        cur_str = replace(cur_str,"''","'")
    End If

    f_DB_IN_STR = cur_str
End Function

'---------------------------------------------------------------------------
' Function Name : f_autoLink
' Description : ���ڿ��� �ڵ���ũ �ɱ�
'---------------------------------------------------------------------------
Function f_autoLink(ByVal str)
    Dim reg

    Set reg = New RegExp

    reg.pattern = "(\w+):\/\/([a-z0-9\_\-\./~@?=%&\-]+)"
    reg.Global = True
    reg.IgnoreCase = True
    str = reg.Replace(str, "$1://$2")

    f_autoLink = str
End Function

'---------------------------------------------------------------------------
' Function Name : f_isVaildProfileImage
' Description : �ش������� �̹��� �������� üũ
'---------------------------------------------------------------------------
Function f_isVaildProfileImage(ByVal imageName)
    Dim imageExt

    imageExt = LCase(Mid(imageName,InStrRev(imageName,".")+1))

    If imageExt <> "jpg" And imageExt <> "gif" And imageExt <> "jpeg" And imageExt <> "bmp" And imageExt <> "jpe" Then
        f_isVaildProfileImage = False
    Else
        f_isVaildProfileImage = True
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : f_FillChar
' Description : �ʿ��� �ڸ�����ŭ Ư�����ڷ� ä���
' Example : Response.Write FillChar(������,ä�ﰪ,����,�ڸ���)
' output :
'---------------------------------------------------------------------------
Function f_FillChar(strValue, FChar, Direction, strLength)
    Dim tmpStr, i

    For i=1 To strLength
        tmpStr = tmpStr & FChar
    Next

    If Direction="L" or Direction="" Then ' ������
        f_FillChar = Right(tmpStr & strValue, strLength)
    Else
        f_FillChar = Left(strValue & tmpStr, strLength)
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : f_ReplaceStr
' Description : NULL CHECK ���ڿ� ġȯ�Լ�
'---------------------------------------------------------------------------
Function f_ReplaceStr(strText, oldString, newString)
    If NOT IsNull(strText) Then
        f_ReplaceStr = Replace(strText, oldString, newString)
    Else
        f_ReplaceStr = ""
    End If
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_SetApostrophe
'INPUT :
'��ɼ��� : DB �Է�ó��
'---------------------------------------------------------------------------
Function f_SetApostrophe(ByVal strVal)
    If Not IsNull(strVal) Then
        strVal = Replace(strVal, "'", "''")
    End If

    f_SetApostrophe = strVal
End Function

'---------------------------------------------------------------------------
'�Լ��� : f_SetTitSTR
'INPUT :
'��ɼ��� : DB �Է�ó��
'---------------------------------------------------------------------------
Function f_SetTitSTR(ByVal strVal)
    If Not IsNull(strVal) Then
        strVal = Replace(strVal, """", "&quot;")
        strVal = Replace(strVal, "'", "&#39;")
    End If

    f_SetTitSTR = strVal
End Function

'---------------------------------------------------------------------------
' Function Name : f_ConvertSpecialChar
' Description : �ױ׹��ڸ� Ư�����ڷθ� ��ȯ.
'---------------------------------------------------------------------------
Function f_ConvertSpecialChar(ByVal StrValue)
    If StrValue <> "" Then
        StrValue = Replace(StrValue, "&", "&amp;")
        StrValue = Replace(StrValue, "<", "&lt;")
        StrValue = Replace(StrValue, ">", "&gt;")
        StrValue = Replace(StrValue, """", "&#34;")
        StrValue = Replace(StrValue, "'", "&#39;")
        StrValue = Replace(StrValue, "|", "&#124;")
        StrValue = Replace(StrValue, Chr(13)&Chr(10), "<br/>")

        f_ConvertSpecialChar = StrValue
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : f_ReConvertSpecialChar
' Description : Ư�����ڸ� �ױ׹��ڷκ�ȯ.
'---------------------------------------------------------------------------
Function f_ReConvertSpecialChar( ByVal strValue )
    If strValue <> "" Then
        strValue = Replace(strValue, "&amp;", "&")
        strValue = Replace(strValue, "&lt;", "<")
        strValue = Replace(strValue, "&gt;", ">")
        strValue = Replace(strValue, "&#34;", """")
        StrValue = Replace(StrValue, "&#39;", "'")
        strValue = Replace(strValue, "&#124;", "|")
        strValue = Replace(strValue, "<br/>", Chr(13)&Chr(10))

        f_ReConvertSpecialChar = strValue
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : f_strCutToByteNoMark
' ��ɼ��� : str�� intByte ���� ��ŭ �ڸ�
'---------------------------------------------------------------------------
Function f_strCutToByteNoMark(ByVal str, ByVal intByte)
    Dim i, tmpByte, tmpStr, strCut

    tmpByte = 0
    tmpStr = null

    If IsNull(str) Or IsEmpty(str) Or str = "" Then
        f_strCutToByteNoMark = ""
        Exit Function
    End If

    If returnByte(str) > intByte Then
        For i = 1 To returnByte(str)
            strCut = Mid(str, i, 1)
            tmpByte = tmpByte + returnByte(strCut)
            tmpStr = tmpStr & strCut

            If tmpByte >= intByte Then
                f_strCutToByteNoMark = tmpStr
                Exit For
            End If
        Next
    Else
        f_strCutToByteNoMark = str
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_FormatDTime
' ��ɼ��� : yyyymmddhhmm ���ڿ��� yyyy/mm/dd hh:mm ���ڿ��� �ٲ۴�.
'---------------------------------------------------------------------------
Private Function gs_FormatDTime(psDTime)
    If IsNull(psDTime) OR Len(psDTime)=0 Then
        gs_FormatDTime=""
    ElseIf Len(psDTime)=12 Then
        gs_FormatDTime = Left(psDTime,4) & "/" & Mid(psDTime,5,2) & "/" & mid(psDTime,7,2) & " " & mid(psDTime,9,2) & ":" & right(psDTime,2)
    Else
        gs_FormatDTime = psDTime
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_FormatDate
' ��ɼ��� : yyyymmdd ���ڿ��� yyyy/mm/dd ���ڿ��� �ٲ۴�.
'           yyyymm ���ڿ��� yyyy/mm ���ڿ��� �ٲ۴�.
'---------------------------------------------------------------------------
Private Function gs_FormatDate(psDate)
    If IsNull(psDate) OR Len(psDate)=0 Then
        gs_FormatDate=""
    ElseIf Len(psDate)=6 Then
        gs_FormatDate = Left(psDate,4) & "/" & Right(psDate,2)
    ElseIf Len(psDate)=8 Then
        gs_FormatDate = Left(psDate,4) & "/" & Mid(psDate,5,2) & "/" & Right(psDate,2)
    Else
        gs_formatDate = psDate
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_Yymmdd
' ��ɼ��� : yyyy/mm/dd ���ڿ��� yyyymmdd ���ڿ��� �ٲ۴�.
'           yyyy/mm ���ڿ��� yyyymm ���ڿ��� �ٲ۴�.
'---------------------------------------------------------------------------
Private Function gs_Yymmdd(psDate)
    If Len(psDate)=10 Then
        gs_Yymmdd = Left(psDate,4) & Mid(psDate,6,2) & Right(psDate,2)
    ElseIf Len(psDate)=7 Then
        gs_Yymmdd = Left(psDate,4) & Right(psDate,2)
    Else
        gs_Yymmdd = psDate
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_yyyymm
' ��ɼ��� : yyyymm���ڿ��� yyyy/mm���ڿ��� �ٲ۴�.
'---------------------------------------------------------------------------
Private function gs_yyyymm(psdate)
    If len(psdate)=6 Then
       gs_yyyymm = left(psdate,4) & "/" & Mid(psdate,5,2)
    ElseIf len(psdate)=8 Then
       gs_yyyymm = left(psdate,4) & "/" & Mid(psdate, 5,2)
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : FormatDate
' ��ɼ��� : YYYYMMDD ���ڿ��� YYYY-MM-DD ���ڿ��� �ٲ۴�.
'           YYYYMM ���ڿ��� YYYY-MM ���ڿ��� �ٲ۴�.
'           YYYYMMDDhhmm ���ڿ��� YYYY-MM-DD hh:mm�� �ٲ۴�.
'---------------------------------------------------------------------------
Private Function gs_FormatDate(psDate)
    If IsNull(psDate) OR Len(psDate)=0 Then
        FormatDate=""
    ElseIf Len(psDate)=6 Then
        FormatDate = Left(psDate,4) & "-" & Right(psDate,2)
    ElseIf Len(psDate)=8 Then
        FormatDate = Left(psDate,4) & "-" & Mid(psDate,5,2) & "-" & Right(psDate,2)
    ElseIf Len(psDate)=12 Then
        FormatDate = Left(psDate,4) & "-" & Mid(psDate,5,2) & "-" & mid(psDate,7,2) & " " & mid(psDate,9,2) & ":" & right(psDate,2)
    Else
        FormatDate = psDate
    End if
End Function

'---------------------------------------------------------------------------
' Function Name : FormatDateLen
' ��ɼ��� : ��¥���ڿ��� ���ϴ� ����(Ln)��ŭ �������ش�.
'           �������� YYYY-MM-DD
'---------------------------------------------------------------------------
Private Function gs_FormatDateLen(psDate, ln)
    If IsNull(psDate) OR Len(psDate)=0 Then
        FormatDateLen=""
    ElseIf ln="6" and Len(psDate)>=6 Then
        FormatDateLen = Left(psDate,4) & "-" & Mid(psDate,5,2)
    ElseIf ln="8" and Len(psDate)>=8 Then
        FormatDateLen = Left(psDate,4) & "-" & Mid(psDate,5,2) & "-" & Mid(psDate,7,2)
    ElseIf ln="12" and Len(psDate)>=12 Then
        FormatDateLen = Left(psDate,4) & "-" & Mid(psDate,5,2) & "-" & mid(psDate,7,2) & " " & mid(psDate,9,2) & ":" & right(psDate,2)
    Else
        FormatDateLen = psDate
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gDotDateLen
' ��ɼ��� : ��¥���ڿ��� ���ϴ� ����(Ln)��ŭ �������ش�.
'           �������� YYYY.MM.DD
'---------------------------------------------------------------------------
Private Function gDotDateLen(psDate, ln)
    If IsNull(psDate) OR Len(psDate)=0 Then
        gDotDateLen=""
    ElseIf ln="6" and Len(psDate)>=6 Then
        gDotDateLen = Left(psDate,4) & "." & Mid(psDate,5,2)
    ElseIf ln="8" and Len(psDate)>=8 Then
        gDotDateLen = Left(psDate,4) & "." & Mid(psDate,5,2) & "." & Mid(psDate,7,2)
    ElseIf ln="12" and Len(psDate)>=12 Then
        gDotDateLen = Left(psDate,4) & "." & Mid(psDate,5,2) & "." & mid(psDate,7,2) & " " & mid(psDate,9,2) & ":" & right(psDate,2)
    Else
        gDotDateLen = psDate
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_FormatNumber
' ��ɼ��� : ���� 123456789 �� 123,456,789 ���ڿ��� �ٲ۴�. (�����Ҽ���)
'---------------------------------------------------------------------------
Private Function gs_FormatNumber(plNum,piPoint)
    If isnull(plNum) Then
        plNum = 0
    Else
        plNum = ccur(plNum)
    End If

    If plNum=0 Then
        gs_FormatNumber = 0
    ElseIf Len(plNum)=0 OR IsNull(plNum) Then
        gs_FormatNumber = ""
    Else
        gs_FormatNumber = FormatNumber(plNum,piPoint)
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_FormatNumber1
' ��ɼ��� : ���� 123456789 �� 123,456,789 ���ڿ��� �ٲ۴�. (�����Ҽ���)
'---------------------------------------------------------------------------
Private Function gs_FormatNumber1(plNum)
    If isnull(plNum) Then
        plNum = 0
    Else
        plNum = ccur(plNum)
    End If

    If plNum=0 Then
        gs_FormatNumber1 = 0
    ElseIf Len(plNum)=0 OR IsNull(plNum) Then
        gs_FormatNumber1 = ""
    Else
        dInt = int(plNum)
        gs_FormatNumber1 = FormatNumber(dInt,0)

        If (plNum-dInt) > 0 Then
            gs_FormatNumber1 = gs_FormatNumber1 & cstr(plNum-dInt)
        End If
    End If
End Function

'---------------------------------------------------------------------------
' Function Name : gs_Add0
' ��ɼ��� : 1-9���ڸ� 01-09 ���ڿ��� �ٲ۴�.
'---------------------------------------------------------------------------
Private Function gs_Add0(piNum)
   If piNum < 10 Then
        gs_Add0 = "0" & cstr(piNum)
   Else
        gs_Add0 = cstr(piNum)
   End If
End Function

'---------------------------------------------------------------------------
' Function Name : gf_LeftAtDb
' ��ɼ��� : �߶� ������ ������ ����
' Input : �����ؽ�Ʈ�� ����, ȭ�鿡 ������ ����Byte ��
'---------------------------------------------------------------------------
Function gf_LeftAtDb(szInput,nLen)
   Dim nCnt
   Dim szLeft

   szInput = Trim(szInput)
   If isNull(szInput) or isEmpty(szInput) Then
      gf_LeftAtDb = ""
   Else
      For nCnt = 1 To Len(szInput)
         szLeft = Mid(szInput,1,nCnt)
         If gf_LenAtDb(szLeft) > nLen Then
            szLeft = Mid(szInput,1,nCnt-1)
            szleft = szleft & "..."
            Exit For
         End If
      Next

      gf_LeftAtDb = szLeft
   End If
End Function

'---------------------------------------------------------------------------
' Function Name : gf_LenAtDb
' ��ɼ��� : �ѱ�/������ üũ�ؼ� �ѱ��� 2Byte�� ������ 1Byte�� �����Ѵ�.
' Input :
'---------------------------------------------------------------------------
Function gf_LenAtDb(szAllText)
    Dim nLen
    Dim nCnt
    Dim szEach

    nLen = 0
    szAllText = Trim(szAllText)

    For nCnt = 1 To Len(szAllText)
        szEach = Mid(szAllText,nCnt,1)

        If 0 <= Asc(szEach) And Asc(szEach) <= 255 Then
            nLen = nLen + 1 '�ѱ��� �ƴ� ���
        Else
            nLen = nLen + 2 '�ѱ��� ���
        End If
    Next

    gf_LenAtDb = nLen
End Function

'---------------------------------------------------------------------------
' Function Name : gf_IpCheck
' ��ɼ��� : IP Address �� ��Ȯ�� ����
' Input : IP Adress
' Example : xxx.xxx.xxx.xxx
' Output : "0" - IP�� �Էµ��� ����
'        "1"  - �������� IP
'        ��Ÿ�޼���-������Ȳ�� ���� �޼���
'---------------------------------------------------------------------------
Function gf_IpCheck(IpAddr)
  If (request("IpAddr")) = "" Then
     gf_IpCheck =  "0"
     Exit Function
  End If

  arr_ip = Split(request("IpAddr") , ".")

  If UBound(arr_ip) <> 3 Then
     gf_IpCheck = "IP Address ���Ŀ���"
     Exit Function
  End If

  If arr_ip(0)="" Or arr_ip(1)="" Or arr_ip(2)="" Or arr_ip(3)="" Then
     gf_IpCheck="outOfRange"
     Exit Function
  End If

  If arr_ip(0) < 1 Or arr_ip(0) > 255 Or arr_ip(1) < 1 Or arr_ip(1) > 255 Or arr_ip(2) < 1 Or arr_ip(2) > 255 Or arr_ip(3) < 1 Or arr_ip(3) > 255 Then
     gf_IpCheck = "IP Address�� ���ڹ����� ������ϴ�."
     Exit Function
  End If

  gf_IpCheck = "1"
End Function

'---------------------------------------------------------------------------
' Function Name : gf_insConvStr
' ��ɼ��� : DB �Է¿� text ��ȯ
' Input :
'---------------------------------------------------------------------------
Function gf_insConvStr(CheckValue)
  CheckValue = Replace(CheckValue, "'", "''")
  CheckValue = Replace(CheckValue, chr(34), "&quot;")
  gf_insConvStr = CheckValue
End Function

'---------------------------------------------------------------------------
' Function Name : gf_viewConvStr
' ��ɼ��� : ������ ��½� html Tag ȿ�� ����
' Input :
'---------------------------------------------------------------------------
Function gf_viewConvStr(CheckValue)
  CheckValue = Replace(CheckValue, "<", "&lt;" )
  CheckValue = Replace(CheckValue,  ">", "&gt;")
  CheckValue = Replace(CheckValue,  "|", "&#124;")
'  CheckValue = Replace(CheckValue,  chr(13), "<br>")
  gf_viewConvStr = CheckValue
End Function

'====================================
'AS���� > AS �Ѱ���Ȳ
'���� �ٿ�ε� ����
'====================================
Function SetAsListExcelErrName(errVal)
	Dim errCode, arrStr, errCnt

	errCode = Split(errVal, ",")
	errCnt = UBound(errCode) - 1

	For i = 0 To UBound(errCode)
		arrStr = arrStr & "'" & Trim(errCode(i)) & "'"

		If i <= CInt(errCnt) Then
			arrStr = arrStr & ","
		End If
	Next

	SetAsListExcelErrName = arrStr
End Function

'====================================
'��¥ ��ȯ(DB �Է¿�)
'����ȣ_20210825
'====================================
Function f_FormatDate()
    Dim f_Date, f_Hour, f_Min

    f_Date = FormatDateTime(Now(), 2)
    f_Hour = FormatDateTime(Now(), 4)
    f_Min = Right(Now(), 3)

    f_FormatDate = f_Date & " " & f_Hour & f_Min
End Function

'====================================
'���ϸ� ��ȯ
'����ȣ_20220107
'====================================

Function f_LogFilename(filename)
    Dim logArr

    logArr = Split(filename, "/")

    f_LogFilename = logArr(UBound(logArr))
End Function
%>