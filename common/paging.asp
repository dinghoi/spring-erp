<%
'#################################################################################
'# 페이징 관련 함수
'#################################################################################

' 페이징시 필요한 전역변수 2개
Dim G_PAGE_SIZE : G_PAGE_SIZE = 10 ' 뿌려질 레코드 개수
Dim G_TOTAL_RECORD ' 전체 레코드 수

'/* 쿼리문 + 페이징 */
Public Function ExecutePage(ByVal pSql, ByVal pPage)
    Dim rs : Set rs = Server.CreateObject("ADODB.RecordSet")
    Dim strSQL
    Dim nPage
    Dim cut, l_sql, r_sql

    If pPage = "" or isNull(pPage) Then pPage = 1

    pSql = UCase(pSql) ' 대문자로 변환
    pSql = Replace(pSql, vbTab, " ") ' 쿼리문의 Tab은 Space로
    pSql = Replace(pSql, vbCr, " ") ' 쿼리문의 개행은 Space로

    cut = InStr(1, pSql, " TOP ")
    ' top 절이 없으면 order by에서 오류 - top절 있는지 검사
    If cut = 0 Then
        cut = InStr(1, pSql, " DISTINCT ")
        If cut > 0 Then
         ' distinct 가 있을 경우
         cut = cut + 8
         l_sql = Left(pSql, cut)
         r_sql = Right(pSql, Len(pSql) - cut - 1)

         pSql = l_sql & " TOP 1000000000 " & r_sql
        Else
         ' distinct 가 없을 경우
         cut = InStr(1, pSql, "SELECT ")
         If cut > 0 Then
         cut = cut + 5
         l_sql = Left(pSql, cut)
         r_sql = Right(pSql, Len(pSql) - cut - 1)

         pSql = l_sql & " TOP 1000000000 " & r_sql
         End If
        End If
    End If

    nPage = pPage * CLng(G_PAGE_SIZE)
    strSQL = "Select TOP " & CStr(nPage) & " * From (" & pSql & ") AS _TEMP_PAGE_TABLE"

    G_TOTAL_RECORD = ExecuteCount(pSql) ' 쿼리문의 전체 레코드 수 구함

    rs.CursorType = 1
    rs.PageSize = G_PAGE_SIZE
    rs.Open strSQL, DBConn

    If Not (rs.EOF Or rs.BOF) Then rs.AbsolutePage = pPage

    Set ExecutePage = rs
End Function

'/* 쿼리문을 주면 그 쿼리문의 레코드 수를 반환 */
Public Function ExecuteCount(ByVal pSql)
    Dim rs : Set rs = Server.CreateObject("ADODB.RecordSet")


    pSql = "Select Count(*) From (" & pSql & ") AS _TEMP_PAGE_TABLE_CNT"
    Set rs = DBConn.Execute(pSql)

    ExecuteCount = CLng(rs(0))
    rs.Close
    Set rs = nothing
End Function

'/* 페이징시 전체 페이지수 계산 */
Function GetPageCount(ByVal pTotalRecord)
    Dim retVal

    pTotalRecord = CLng(pTotalRecord)
    retVal = Fix(pTotalRecord / G_PAGE_SIZE)
    If (pTotalRecord Mod G_PAGE_SIZE) > 0 Then
        retVal = retVal + 1
    End If
    GetPageCount = CLng(retVal)
End Function

'/* 페이지 네비게이션을 뿌려주는 함수 */
Public Function ShowPageBar(ByVal pCurPage, ByVal pPreImg, ByVal pNextImg, ByVal param)
    Dim nPREV
    Dim nCUR
    Dim nNEXT
    Dim i
    Dim nPageCount
    Dim retVal
    Dim strLink
    Dim pageKubun


    If pCurPage = "" or isNull(pCurPage) Then pCurPage = 1

    nPageCount = GetPageCount(G_TOTAL_RECORD)

    If pPreImg = "" Then
        pPreImg = "[이전]"
    Else
        pPreImg = "<img src='" & pPreImg & "' border=0 align=absmiddle>"
    End If

    If pNextImg = "" Then
        pNextImg = "[다음]"
    Else
        pNextImg = "<img src='" & pNextImg & "' border=0 align=absmiddle>"
    End If

    nPREV = (Fix((pCurPage - 1) / 10) - 1) * 10 + 1
    nCUR = (Fix((pCurPage - 1) / 10)) * 10 + 1
    nNEXT = (Fix((pCurPage - 1) / 10) + 1) * 10 + 1


    ' [이전] 페이지 조합
    If nPREV > 0 Then
        strLink = "?curPage=" & nPREV & param
        retVal = "<a href=""" & strLink & """>" & pPreImg & "</a> "
    Else
        retVal = "" & pPreImg & " "
    End If
    i = 1
    Do While i < 11 And nCUR <= nPageCount
        If nCUR = nPageCount Or i = 10 Then
         pageKubun = " "
        Else
         pageKubun = " . "
        End If

        If CInt(pCurPage) = CInt(nCUR) Then
         retVal = retVal & "<font color=#FF6700 size=3><b>" & nCUR & "</b></font>" & pageKubun
        Else
         strLink = "?curPage=" & nCUR & param
         retVal = retVal & "<a href=""" & strLink & """>" & nCUR & "</a>" & pageKubun
        End If
        nCUR = nCUR + 1
        i = i + 1
    Loop
    ' [다음] 페이지 조합
    If nNEXT <= nPageCount Then
        strLink = "?curPage=" & nNEXT & param
        retVal = retVal & " <a href=""" & strLink & """>" & pNextImg & "</a>"
    Else
        retVal = retVal & pNextImg & ""
    End If

    ShowPageBar = retVal
End Function
%>