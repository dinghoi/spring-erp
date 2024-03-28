<%
'====================================
'Author : 허정호
'Modify Date : 20201117
'Desc : 공통 사용 변수 정의 주가
'	파일 include 시 "/include/nkpmg_user.asp" 파일이 먼저 include 되어야 함
'====================================

'====================================
' 공용 사용 변수
'====================================
Dim SYSDATE, SYSDATE12, OrderByOrgName

' 시스템 날짜(문자열 8자리, YYYYMMDD )
SYSDATE = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)

' 시스템 날짜(문자열 12자리, YYYYMMDDhhmm )
SYSDATE12 = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)

'시스템 관리자 권한
Dim SysAdminYn

Select Case user_id
	'운영(허정호)
	Case "102592"
		SysAdminYn = "Y"
	Case Else
		SysAdminYn = "N"
End Select

'조직명 정렬 우선 기준
OrderByOrgName = "'케이원', '케이네트웍스', '케이시스템', '엔와이컴퓨터', '에스유에이치', '케이더봄', '엔와이리테일'"
'====================================
'개인 업무 관리 메뉴 접속 권한 관리
'====================================
Dim IntroMemberYn, MemberPayAdmin, u_id, m_seq, m_name

u_id = f_toString(Request.cookies("nkp_member")("coo_user_id"), "")
m_seq = f_toString(Request.cookies("nkp_member")("coo_m_seq"), "")
m_name = f_toString(Request.cookies("nkp_member")("coo_m_name"), "")

IntroMemberYn = "N"'신규 채용 여부

If u_id = "kone" Then
	IntroMemberYn = "Y"
End If

MemberPayAdmin = "N"'직원현황 열람 권한

If position = "총괄대표" Or position = "본부장" Or position = "사업부장" Or SysAdminYn = "Y" Then
	MemberPayAdmin = "Y"
End If
'====================================
'서비스 관리 메뉴 접속 권한 관리
'====================================

'====================================
'인사 관리 메뉴 접속 권한
'====================================
Dim InsaMasterModYn	'인사 기본 정보 수정 권한 여부
Dim InsaMasterDelYn	'인사 기본 정보 삭제 권한 여부
Dim InsaCarDelYn	'차량 정보 삭제 권한

'인사 기본 정보 수정 권한
Select Case user_id
	'운영(송지영, 지현주, 이윤영, 김정훈, 허정호, 노혜진)
	Case "100018", "101622", "100104", "102560", "102592", "102615"
		InsaMasterModYn = "Y"
	Case Else
		InsaMasterModYn = "N"
End Select

'인사 기본 정보 삭제 권한
Select Case user_id
	'운영(송지영, 지현주, 허정호, 노혜진)
	Case "100018", "101622", "102592", "102615"
		InsaMasterDelYn = "Y"
	Case Else
		InsaMasterDelYn = "N"
End Select

'차량 정보 삭제 권한
Select Case user_id
	'(이윤영, 허정호)
	Case "100104", "102592"
		InsaCarDelYn = "Y"
	Case Else
		InsaCarDelYn = "N"
End Select

'====================================
'급여 관리 메뉴 접속 권한
'====================================

'====================================
'비용 관리 메뉴 접속 권한
'====================================
Dim resideEndViewYn

'비용 마감 > 상주직접비/공통비 항목 노출 권한 설정
Select Case user_id
	'Case "900001", "100359", "100952", "101100", "102305", "102306", "102592"
	Case "100359", "102592"
		resideEndViewYn = "Y"
	Case Else
		resideEndViewYn = "N"
End Select

'====================================
'영업 관리 메뉴 접속 권한
'====================================
Dim ProfitLossYn
Dim SubEmpReportYn
Dim SubProfitLossYn
Dim SubKdcEmpReportYn

'손익현황 메뉴 접속 권한 여부
'If CInt(coo_sales_grade) < 2 Or (coo_bonbu = "ITO 사업본부" And coo_position = "사업부장") Or (coo_bonbu = "ITO 사업본부" And coo_position = "본부장") Then
'	ProfitLossYn = "Y"
'Else
'	ProfitLossYn = "N"
'End If

If ProfitLossYn = "Y" Then
	'사업부별 인원 현황 서브 메뉴 접속 권한 여부
	Select Case coo_user_id
		Case "90001"
			SubEmpReportYn = "Y"
		Case Else
			SubEmpReportYn = "N"
	End Select

	'사업부별 월별손익 서브 메뉴 접속 권한 여부
	Select Case coo_user_id
		'김희찬, 조대현, 차재명, 이홍석, 강경진, 최성민(102305, 102306)
		Case "100952", "100703", "101100", "101664", "101880", "102305", "102306", "ktrental2"
			SubProfitLossYn = "Y"
		Case Else
			SubProfitLossYn = "N"
	End Select

	'사업부별 인원 현황(kdc) 서브 메뉴 접속 권한 여부
	Select Case coo_user_id
		Case "100703", "101100", "101664", "101880", "102305", "102306", "ktrental2"
			SubKdcEmpReportYn = "Y"
		Case Else
			SubKdcEmpReportYn = "N"
	End Select
End If

'사업부별 월별손익(수정) 메뉴 접속 권한 설정[허정호_20201221]
Dim ProfitLossMonthNewYn, CompanyCostYn, CoworkYn

Select Case user_id
	'강경진
	Case "101880"
		ProfitLossMonthNewYn = "Y"
	Case Else
		ProfitLossMonthNewYn = "N"
End Select

'거래처 손익 View 권한
Select Case user_id
	'대표이사, 사장
	Case "100001", "100740"
		CompanyCostYn = "Y"

	'부사장, 대표이사(케이더봄)
	'Case "101245", "100262", "102627"
	''	CompanyCostYn = "Y"
	'본부장
	'Case "102663"
	''	CompanyCostYn = "Y"
	'팀장
	'Case "100020", "100125", "100173", "100180", "100186", "100187", "100244", "100703", "100953", "101246", "102271", "102665"
	''	CompanyCostYn = "Y"

	'재무이사
	Case "100359"
		CompanyCostYn = "Y"
	Case Else
		CompanyCostYn = "N"
End Select

'협업 View 권한
Select Case user_id
	'대표이사, 사장
	Case "100001", "100740"
		CoworkYn = "Y"
	'부사장, 본부장
	'Case "101245", "100262", "102663"
	'	CoworkYn = "Y"
	'팀장
	'Case "100020", "100125", "100173", "100180", "100186", "100187", "100244", "100703", "100953", "101246", "102271", "102665"
	'	CoworkYn = "Y"
	'재무이사
	Case "100359"
		CoworkYn = "Y"
	Case Else
		CoworkYn = "N"
End Select

'사업부별 인원현황(KDC) 메뉴 접속 권한 설정[허정호_20201221]
Dim empReportKDCYn

Select Case user_id
	'조대현
	Case "100703"
		empReportKDCYn = "Y"
	Case Else
		empReportKDCYn = "N"
End Select

'사업부별 손익총괄 상세 링크 접근 권한(수정)
Dim empProfitViewAll, empProfitViewSI, empProfitViewNI, empProfitGrade
empProfitViewAll = "N"
empProfitViewSI = "N"
empProfitViewNI = "N"
empProfitGrade = "N"	'본부장 권한

Select Case user_id
	'김승일(회장), 송관섭(대표), 박정신(재무이사), 허정호(시스템관리자), 김찬양(재무), 박성민(재무), 박명호(재무), 박찬규(재무), 손기봉(고문)
	Case "100001", "100740", "100359", "102592", "100842", "101672", "101902", "102825", "102498"
		'전사 노출
		empProfitViewAll = "Y"
		empProfitGrade = "Y"
	'권회철
	Case "101245"
		'SI1, SI2 노출
		empProfitViewSI = "Y"
		empProfitGrade = "Y"
	'이동규
	Case "102652"
		'NI, ICT 노출
		empProfitViewNI = "Y"
		empProfitGrade = "Y"
End Select

'추가 본부장 권한
Select Case user_id
	'이재준, 강명석, 최길성, 여우진, 고대준, 전대석, 조대현, 유상훈
	Case "102489", "100125", "100031", "100021", "101246", "100262", "100703", "102926"
		empProfitGrade = "Y"
End Select

'부문공통비, 부문공통비(2) 권한
Dim partCostView, subPartCostView

Select Case bonbu
	Case "SI1본부", "SI2본부", "NI본부", "공공본부"
		partCostView = "Y"
	Case "금융SI본부", "공공SI본부", "DI사업부문"
		subPartCostView = "Y"
	Case Else
		partCostView = "N"
		subPartCostView = "N"
End Select

'====================================
'상품 자재 관리 메뉴 접속 권한
'====================================
Dim NwInMenuYn

'N/W 입출고 접속 권한 설정

Select Case user_id
	'박정신, 전간수, 허정호
	Case "100359", "100015", "102592"
		NwInMenuYn = "Y"
	Case Else
		NwInMenuYn = "N"
End Select



'====================================
'회계 관리 메뉴 접속 권한
'====================================

'====================================
'임원 정보 관리 메뉴 접속 권한
'====================================
%>
