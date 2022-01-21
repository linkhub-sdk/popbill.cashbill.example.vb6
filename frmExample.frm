VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 현금영수증 SDK 예제"
   ClientHeight    =   11175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16455
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   16455
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12120
      TabIndex        =   70
      Top             =   240
      Width           =   3975
   End
   Begin VB.Frame Frame7 
      Caption         =   "현금영수증 관련 기능"
      Height          =   7185
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   15570
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스 "
         Height          =   2415
         Left            =   1800
         TabIndex        =   46
         Top             =   1440
         Width           =   3375
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "삭제"
            Height          =   375
            Left            =   1920
            Style           =   1  '그래픽
            TabIndex        =   49
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_sub 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   48
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   430
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.Line Line5 
            X1              =   960
            X2              =   2370
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   480
            Top             =   360
            Width           =   2415
         End
         Begin VB.Line Line4 
            X1              =   960
            X2              =   960
            Y1              =   1680
            Y2              =   840
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " 인쇄/보기"
         Height          =   2760
         Left            =   7680
         TabIndex        =   32
         Top             =   4125
         Width           =   5370
         Begin VB.CommandButton btnGetViewURl 
            Caption         =   "현금영수증 보기 URL(메뉴x)"
            Height          =   375
            Left            =   210
            TabIndex        =   65
            Top             =   840
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPDFURL 
            Caption         =   "PDF 다운로드 URL"
            Height          =   390
            Left            =   3000
            TabIndex        =   63
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "공급받는자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1740
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "현금영수증 보기 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   390
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "현금영수증 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   1305
            Width           =   2625
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "대량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   2190
            Width           =   2625
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "현금영수증 메일링크 URL"
            Height          =   390
            Left            =   3000
            TabIndex        =   33
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   1890
         Left            =   13200
         TabIndex        =   28
         Top             =   4125
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   31
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "발행 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   30
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "매출 작성"
            Height          =   390
            Left            =   210
            TabIndex        =   29
            Top             =   1275
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "부가 기능"
         Height          =   2775
         Left            =   2760
         TabIndex        =   24
         Top             =   4125
         Width           =   4815
         Begin VB.CommandButton btnAssignMgtKey 
            Caption         =   "문서번호 할당"
            Height          =   390
            Left            =   2760
            TabIndex        =   64
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnUpdateemailconfig 
            Caption         =   "알림메일 전송설정 수정"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   2280
            Width           =   2415
         End
         Begin VB.CommandButton btnListemailconfig 
            Caption         =   "알림메일 전송목록 조회"
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   27
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   26
            Top             =   825
            Width           =   2415
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   390
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " 문서 정보 "
         Height          =   2775
         Left            =   240
         TabIndex        =   19
         Top             =   4125
         Width           =   2265
         Begin VB.CommandButton btnSearch 
            Caption         =   "목록 조회"
            Height          =   390
            Left            =   195
            TabIndex        =   50
            Top             =   1800
            Width           =   1845
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "상세 정보확인"
            Height          =   390
            Left            =   195
            TabIndex        =   23
            Top             =   1320
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "상태 변경이력"
            Height          =   390
            Left            =   195
            TabIndex        =   22
            Top             =   2280
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "상태 대량 확인"
            Height          =   390
            Left            =   195
            TabIndex        =   21
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "상태 확인"
            Height          =   390
            Left            =   195
            TabIndex        =   20
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "취소현금영수증 즉시발행 프로세스"
         Height          =   2385
         Left            =   5640
         TabIndex        =   17
         Top             =   1440
         Width           =   4095
         Begin VB.CommandButton btnRegistIssue_part 
            Caption         =   "부분취소 즉시발행"
            Height          =   375
            Left            =   1680
            TabIndex        =   60
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton btnRevokeRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   375
            Left            =   480
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   52
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   480
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   1560
            Width           =   960
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "삭제"
            Height          =   375
            Left            =   1920
            Style           =   1  '그래픽
            TabIndex        =   18
            Top             =   1560
            Width           =   855
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   3495
         End
         Begin VB.Line Line1 
            X1              =   960
            X2              =   960
            Y1              =   1905
            Y2              =   600
         End
         Begin VB.Line Line2 
            X1              =   1080
            X2              =   2490
            Y1              =   1760
            Y2              =   1760
         End
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "문서번호 사용여부 확인"
         Height          =   375
         Left            =   6150
         TabIndex        =   16
         Top             =   255
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   3300
         TabIndex        =   15
         Top             =   285
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "국세청에 신고된 현금영수증을 취소하기 위해서는'취소현금영수증'을 발행해야 합니다."
         Height          =   375
         Left            =   5640
         TabIndex        =   53
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "문서번호( MgtKey) : "
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   15810
      Begin VB.Frame Frame15 
         Caption         =   "파트너과금 포인트"
         Height          =   2295
         Left            =   13200
         TabIndex        =   56
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "연동과금 포인트"
         Height          =   2295
         Left            =   10920
         TabIndex        =   54
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련"
         Height          =   2295
         Left            =   8880
         TabIndex        =   43
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   2295
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   2295
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   2295
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL "
         Height          =   2295
         Left            =   6600
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   210
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   225
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   11400
      TabIndex        =   69
      Top             =   285
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4320
      TabIndex        =   2
      Top             =   285
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' 팝빌 현금영수증 API VB 6.0 SDK Example
'
' - 업데이트 일자 : 2022-01-17
' - 연동 기술지원 연락처 : 1600-9854
' - 연동 기술지원 이메일 : code@linkhubcorp.com
' - VB6 SDK 연동환경 설정방법 안내 : https://docs.popbill.com/cashbill/tutorial/vb
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
'=========================================================================


Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'현금영수증 서비스 객체 생성
Private CashbillService As New PBCBService

'=========================================================================
' 팝빌 사이트를 통해 발행하여 문서번호가 부여되지 않은 현금영수증에 문서번호를 할당합니다.
' - https://docs.popbill.com/cashbill/vb/api#AssignMgtKey
'=========================================================================
Private Sub btnAssignMgtKey_Click()
    Dim Response As PBResponse
    Dim itemKey As String
    Dim mgtKey As String
    
    '현금영수증 아이템키, 목록조회(Search) API의 반환항목중 ItemKey 참조
    itemKey = "020042413523200001"
            
    '할당할 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    mgtKey = "20220101-05"
        
    Set Response = CashbillService.AssignMgtKey(txtCorpNum.Text, itemKey, mgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/cashbill/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 PDF 파일을 다운 받을 수 있는 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
'=========================================================================
Private Sub btnGetPDFURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPDFURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub


'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/cashbill/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 발행시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = CashbillService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "발행단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 팝빌 현금영수증 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = CashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim URL As String
     
    URL = CashbillService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/cashbill/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "vb6Cashbill001"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "qwe123!@#"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
    
        
    Set Response = CashbillService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' https://docs.popbill.com/cashbill/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    '확인할 담당자 아이디
    ContactID = "testkorea"
    
    Set info = CashbillService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = CashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/cashbill/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = "vb6Cashbill001"
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
    
                
    Set Response = CashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = CashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자 성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
' - https://docs.popbill.com/cashbill/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = CashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = CashbillService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = CashbillService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/cashbill/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim URL As String
     
    URL = CashbillService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/cashbill/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = CashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 파트너가 현금영수증 관리 목적으로 할당하는 문서번호 사용여부를 확인합니다.
' - 이미 사용 중인 문서번호는 중복 사용이 불가하고, 현금영수증이 삭제된 경우에만 문서번호의 재사용이 가능합니다.
' - 문서번호는 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
' - https://docs.popbill.com/cashbill/vb/api#CheckMgtKeyInUse
'=========================================================================
Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.CheckMgtKeyInUse(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 작성된 현금영수증 데이터를 팝빌에 저장과 동시에 발행하여 "발행완료" 상태로 처리합니다.
' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=vb
' - https://docs.popbill.com/cashbill/vb/api#RegistIssue
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Cashbill As New PBCashbill
    Dim Response As PBResponse
    Dim emailSubject As String
    
    '현금영수증 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    Cashbill.mgtKey = txtMgtKey.Text
    
    '[취소거래시 필수] 원본 국세청승인번호
    '문서정보(GetInfo API)의 응답항목중 국세청승인번호(confirmNum)를 확인하여 기재
    Cashbill.orgConfirmNum = ""
    
    '[취소거래시 필수] 원본 거래일자
    '문서정보(GetInfo API)의 응답항목중 거래일자(tradeDate)를 확인하여 기재
    Cashbill.orgTradeDate = ""
    
    '문서형태, [승인거래, 취소거래] 중 기재
    Cashbill.tradeType = "승인거래"
    
    '거래구분, [소득공제용, 지출증빙용] 중 기재
    Cashbill.tradeUsage = "소득공제용"
    
    '거래유형, [일반, 도서공연, 대중교통] 중 기재
    Cashbill.tradeOpt = "일반"
    
    '과세형태, [과세, 비과세] 중 기재
    Cashbill.taxationType = "과세"
    
    '거래금액, 공급가액 + 봉사료 + 세액
    Cashbill.totalAmount = "11000"
    
    '공급가액
    Cashbill.supplyCost = "10000"
    
    '부가세
    Cashbill.tax = "1000"
    
    '봉사료
    Cashbill.serviceFee = "0"
    
    '가맹점 사업자번호, "-" 제외 10자리
    Cashbill.franchiseCorpNum = "1234567890"
    
    '가맹점 종사업장 식별번호
    Cashbill.franchiseTaxRegID = ""
    
    '가맹점 상호
    Cashbill.franchiseCorpName = "발행자 상호"
    
    '가맹점 대표자 성명
    Cashbill.franchiseCEOName = "발행자 대표자"
    
    '가맹점 주소
    Cashbill.franchiseAddr = "발행자 주소"
    
    '가맹점 전화번호
    Cashbill.franchiseTEL = "070-1234-1234"
        
    '식별번호, 거래구분에 따라 작성
    '소득공제용 - 주민등록/휴대폰/카드번호 기재가능
    '지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호 기재가능
    Cashbill.identityNum = "0101112222"
        
    '주문자명
    Cashbill.customerName = "주문자명"
    
    '주문상품명
    Cashbill.itemName = "주문상품명"
    
    '주문번호
    Cashbill.orderNumber = "주문번호"
    
    '주문자 이메일
    '팝빌 개발환경에서 테스트하는 경우에도 안내 메일이 전송되므로,
    '실제 거래처의 메일주소가 기재되지 않도록 주의
    Cashbill.email = "test@test.com"
    
    '주문자 휴대폰
    Cashbill.hp = "010-111-222"
    
    '현금영수증 발행 알림문자 전송여부
    Cashbill.smssendYN = False
    
    '안내메일 제목, 미기재시 기본양식으로 전송.
    emailSubject = ""
            
    Set Response = CashbillService.RegistIssue(txtCorpNum.Text, Cashbill, txtUserID.Text, emailSubject)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message + vbCrLf + "국세청 승인번호 : " + Response.confirmNum + vbCrLf + "거래일자 : " + Response.tradeDate)
End Sub

'=========================================================================
' 국세청 전송 이전 "발행완료" 상태의 현금영수증을 "발행취소"하고 국세청 신고 대상에서 제외합니다.
' - Delete(삭제)함수를 호출하여 "발행취소" 상태의 현금영수증을 삭제하면, 문서번호 재사용이 가능합니다.
' - https://docs.popbill.com/cashbill/vb/api#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_sub_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "발행취소 메모"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 삭제 가능한 상태의 현금영수증을 삭제합니다.
' - 삭제 가능한 상태: "임시저장", "발행취소", "전송실패"
' - 현금영수증을 삭제하면 사용된 문서번호(mgtKey)를 재사용할 수 있습니다.
' - https://docs.popbill.com/cashbill/vb/api#Delete
'=========================================================================
Private Sub btnDelete_sub_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 취소 현금영수증 데이터를 팝빌에 저장과 동시에 발행하여 "발행완료" 상태로 처리합니다.
' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=vb
' - https://docs.popbill.com/cashbill/vb/api#RevokeRegistIssue
'=========================================================================
Private Sub btnRevokeRegistIssue_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    Dim smssendYN As Boolean
    Dim memo As String
    
    '원본현금영수증 승인번호
    orgConfirmNum = "TB0000037"
    
    '원본현금영수증 거래일자
    orgTradeDate = "20220101"
    
    '발행안내문자 전송여부
    smssendYN = False
    
    '메모
    memo = "취소현금영수증 즉시발행"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message + vbCrLf + "국세청 승인번호 : " + Response.confirmNum + vbCrLf + "거래일자 : " + Response.tradeDate)
End Sub

'=========================================================================
' 1건의 취소 현금영수증 데이터를 팝빌에 저장과 동시에 발행하여 "발행완료" 상태로 처리합니다.
' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=dotnet
' - https://docs.popbill.com/cashbill/vb/api#RevokeRegistIssue
'=========================================================================
Private Sub btnRegistIssue_part_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    Dim smssendYN As Boolean
    Dim memo As String
    Dim isPartCancel As Boolean
    Dim cancelType As Integer
    Dim supplyCost As String
    Dim tax As String
    Dim serviceFee As String
    Dim totalAmount As String
    
    '원본현금영수증 승인번호
    orgConfirmNum = "TB0000037"
    
    '원본현금영수증 거래일자
    orgTradeDate = "20220101"
    
    '발행안내문자 전송여부
    smssendYN = False
    
    '메모
    memo = "취소현금영수증 즉시발행"
    
    '부분취소 여부, True-부분취소/False-전체취소
    isPartCancel = True
    
    '취소사유, 1-거래취소, 2-오류발급취소, 3-기타
    cancelType = 1
    
    '[취소] 공급가액
    supplyCost = "7000"
    
    '[취소] 세액
    tax = "700"
    
    '[취소] 봉사료
    serviceFee = "0"
    
    '[취소] 합계금액
    totalAmount = "7700"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message + vbCrLf + "국세청 승인번호 : " + Response.confirmNum + vbCrLf + "거래일자 : " + Response.tradeDate)
End Sub

'=========================================================================
' 국세청 전송 이전 "발행완료" 상태의 현금영수증을 "발행취소"하고 국세청 신고 대상에서 제외합니다.
' - Delete(삭제)함수를 호출하여 "발행취소" 상태의 현금영수증을 삭제하면, 문서번호 재사용이 가능합니다.
' - https://docs.popbill.com/cashbill/vb/api#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 삭제 가능한 상태의 현금영수증을 삭제합니다.
' - 삭제 가능한 상태: "임시저장", "발행취소", "전송실패"
' - 현금영수증을 삭제하면 사용된 문서번호(mgtKey)를 재사용할 수 있습니다.
' - https://docs.popbill.com/cashbill/vb/api#Delete
'=========================================================================
Private Sub btnDelete_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 1건의 상태 및 요약정보를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetInfo
'=========================================================================
Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
    Dim tmp As String
    
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey (팝빌번호) : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey (문서번호) : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate (거래일자) : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeType (문서형태) : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "tradeUsage (거래구분) : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeOpt (거래유형) : " + cbInfo.tradeOpt + vbCrLf
    tmp = tmp + "taxationType (과세형태) : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount (거래금액) : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "issueDT (발행일시) : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "regDT (등록일시) : " + cbInfo.regDT + vbCrLf
    tmp = tmp + "stateMemo (상태메모) : " + cbInfo.stateMemo + vbCrLf
    tmp = tmp + "stateCode (상태코드) : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (상태변경일시) : " + cbInfo.stateDT + vbCrLf
    tmp = tmp + "identityNum (식별번호) : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "itemName (주문상품명) : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "customerName (주문자명) : " + cbInfo.customerName + vbCrLf
    tmp = tmp + "confirmNum (국세청승인번호) : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "orgConfirmNum (원본 현금영수증 국세청승인번호) : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (원본 현금영수증 거래일자) : " + cbInfo.orgTradeDate + vbCrLf
    tmp = tmp + "ntssendDT (국세청 전송일시) : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (국세청 처리결과 수신일시) : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode (국세청 처리결과 상태코드) : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage (국세청 처리결과 메시지) : " + cbInfo.ntsresultMessage + vbCrLf
    tmp = tmp + "printYN (인쇄여부) : " + CStr(cbInfo.printYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 다수건의 현금영수증 상태 및 요약 정보를 확인합니다. (1회 호출 시 최대 1,000건 확인 가능)
' - https://docs.popbill.com/cashbill/vb/api#GetInfos
'=========================================================================
Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBCbInfo
    
    '현금영수증 문서번호배열, 최대 1000건
    KeyList.Add "20220101-001"
    KeyList.Add "20220101-002"
    KeyList.Add "20220101-003"
    KeyList.Add "20220101-004"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "itemKey (팝빌번호) | mgtKey (문서번호) | tradeDate (거래일자) | tradeType (문서형태) | tradeUsage (거래구분) | tradeOpt (거래유형) |  " + _
          "taxationType (과세형태) | totalAmount (거래금액) | issueDT (발행일시) | regDT (등록일시) | stateMemo (상태메모) | stateCode (상태코드)  " + _
          "stateDT (상태변경일시) | identityNum (식별번호) | itemName (주문상품명) | customerName (주문자명) | confirmNum (국세청승인번호)  " + _
          "orgConfirmNum (원본 현금영수증 국세청승인번호) | orgTradeDate (원본 현금영수증 거래일자) | ntssendDT (국세청 전송일시)  " + _
          "ntsresultDT (국세청 처리결과 수신일시) | ntsresultCode (국세청 처리결과 상태코드) | ntsresultMessage (국세청 처리결과 메시지) | printYN (인쇄여부) " + vbCrLf + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + info.mgtKey + " | " + info.tradeDate + " | " + info.tradeType + " | " + info.tradeUsage + " | " + info.tradeOpt + " | " + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | " + info.issueDT + " | " + info.regDT + " | " + info.stateMemo + " | " + CStr(info.stateCode) + " | " + info.stateDT + " | " + info.identityNum + " | "
        tmp = tmp + info.itemName + " | " + info.customerName + " | " + info.confirmNum + " | " + info.orgConfirmNum + " | " + info.orgTradeDate + " | " + info.ntssendDT + " | " + info.ntsresultDT + " | "
        tmp = tmp + info.ntsresultCode + " | " + info.ntsresultMessage + " | " + CStr(info.printYN) + vbCrLf
                
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 현금영수증 1건의 상세정보를 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetDetailInfo
'=========================================================================
Private Sub btnGetDetailInfo_Click()
    Dim cbDetailInfo As PBCashbill
    Dim tmp As String
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "mgtKey (문서번호) : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "confirmNum (국세청승인번호) : " + cbDetailInfo.confirmNum + vbCrLf
    tmp = tmp + "tradeDate (거래일자) : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage (거래구분) : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeOpt (거래유형) : " + cbDetailInfo.tradeOpt + vbCrLf
    tmp = tmp + "tradeType (문서형태) : " + cbDetailInfo.tradeType + vbCrLf
    tmp = tmp + "taxationType (과세형태) : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost (공급가액) : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax (부가세) : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee (봉사료) : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount (거래금액) : " + cbDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "orgConfirmNum (원본현금영수증 국세청승인번호) : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (원본현금영수증 거래일자) : " + cbDetailInfo.orgTradeDate + vbCrLf
    tmp = tmp + "cancelType (취소사유) : " + CStr(cbDetailInfo.cancelType) + vbCrLf
    tmp = tmp + "franchiseCorpNum (가맹점 사업자번호) : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseTaxRegID (가맹점 종사업장 식별번호) : " + cbDetailInfo.franchiseTaxRegID + vbCrLf
    tmp = tmp + "franchiseCorpName (가맹점 상호) : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName (가맹점 대표자 성명) : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr (가맹점 주소) : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL (가맹점 전화번호) : " + cbDetailInfo.franchiseTEL + vbCrLf
    tmp = tmp + "identityNum (식별번호) : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName (주문자명) : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName (주문상품명) : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber (주문번호) : " + cbDetailInfo.orderNumber + vbCrLf
    tmp = tmp + "email (주문자 이메일) : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp (주문자 휴대폰) : " + cbDetailInfo.hp + vbCrLf
    
    tmp = tmp + "smssendYN (발행안내문자 전송여부) : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 검색조건에 해당하는 현금영수증을 조회합니다. (조회기간 단위 : 최대 6개월)
' - https://docs.popbill.com/cashbill/vb/api#Search
'=========================================================================
Private Sub btnSearch_Click()
    Dim cbSearchList As PBCBSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim tradeOpt As New Collection
    Dim taxationType As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim QString As String
    Dim franchiseTaxRedID As String
    
    '[필수] 일자유형, R-등록일자, T-거래일자 I-발행일시
    DType = "T"
    
    '[필수] 시작일자, 형식(yyyyMMdd)
    SDate = "20220101"
    
    '[필수] 종료일자, 형식(yyyyMMdd)
    EDate = "20220130"
    
    '상태코드 배열, 미기재시 전체 상태조회, 상태코드(stateCode)값 3자리의 배열, 2,3번째 자리에 와일드카드 가능
    state.Add "3**"
    state.Add "4**"
    
    '문서형태 배열, N-일반 현금영수증, C-취소 현금영수증
    tradeType.Add "N"
    tradeType.Add "C"
        
    '거래구분 배열, P-소득공제용, C-지출증빙용
    tradeUsage.Add "P"
    tradeUsage.Add "C"
    
    '거래유형 배열, N-일반, B-도서공연, T-대중교통
    tradeOpt.Add "N"
    tradeOpt.Add "B"
    tradeOpt.Add "T"
    
    '과세형태 배열, T-과세, N-비과세
    taxationType.Add "T"
    taxationType.Add "N"
                
    '페이지 번호, 기본값 1
    Page = 1
    
    '페이지당 목록갯수, 기본값 500
    PerPage = 30
    
    '정렬방향 D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '현금영수증 식별번호 조회, 미기재시 전체조회
    QString = ""
    
    '가맹점 종사업장 번호, 미기재시 전체조회
    '└ 다수건 검색시 콤마(",")로 구분. 예) 1234,1000
    franchiseTaxRedID = ""
    
    
    Set cbSearchList = CashbillService.Search(txtCorpNum.Text, DType, SDate, EDate, state, tradeType, tradeUsage, taxationType, Page, PerPage, Order, QString, tradeOpt, franchiseTaxRedID)
     
    If cbSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (응답코드) : " + CStr(cbSearchList.code) + vbCrLf
    tmp = tmp + "total (검색결과 건수) : " + CStr(cbSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(cbSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(cbSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(cbSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + cbSearchList.message + vbCrLf + vbCrLf + vbCrLf
    
    tmp = "itemKey (팝빌번호) | mgtKey (문서번호) | tradeDate (거래일자) | tradeType (문서형태) | tradeUsage (거래구분) | tradeOpt (거래유형) |  " + _
         "taxationType (과세형태) | totalAmount (거래금액) | issueDT (발행일시) | regDT (등록일시) | stateMemo (상태메모) | stateCode (상태코드)  " + _
         "stateDT (상태변경일시) | identityNum (식별번호) | itemName (주문상품명) | customerName (주문자명) | confirmNum (국세청승인번호)  " + _
         "orgConfirmNum (원본 현금영수증 국세청승인번호) | orgTradeDate (원본 현금영수증 거래일자) | ntssendDT (국세청 전송일시)  " + _
         "ntsresultDT (국세청 처리결과 수신일시) | ntsresultCode (국세청 처리결과 상태코드) | ntsresultMessage (국세청 처리결과 메시지) | printYN (인쇄여부) " + vbCrLf + vbCrLf
          
    Dim info As PBCbInfo
    
    For Each info In cbSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + info.mgtKey + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.tradeType + " | "
        tmp = tmp + info.tradeUsage + " | "
        tmp = tmp + info.tradeOpt + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | "
        tmp = tmp + info.issueDT + " | "
        tmp = tmp + info.regDT + " | "
        tmp = tmp + info.stateMemo + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.stateDT + " | "
        tmp = tmp + info.identityNum + " | "
        tmp = tmp + info.itemName + " | "
        tmp = tmp + info.customerName + " | "
        tmp = tmp + info.confirmNum + " | "
        tmp = tmp + info.orgConfirmNum + " | "
        tmp = tmp + info.orgTradeDate + " | "
        tmp = tmp + info.ntssendDT + " | "
        tmp = tmp + info.ntsresult + " | "
        tmp = tmp + info.ntsresultDT + " | "
        tmp = tmp + info.ntsresultCode + " | "
        tmp = tmp + info.ntsresultMessage + " | "
        tmp = tmp + CStr(info.printYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 현금영수증의 상태에 대한 변경이력을 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetLogs
'=========================================================================
Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim log As PBCbLog
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType(로그타입) | Log(이력정보) | ProcType(처리형태) | ProcMemo(처리메모) | RegDT(등록일시) | IP(아이피)" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 현금영수증과 관련된 안내 메일을 재전송 합니다.
' - https://docs.popbill.com/cashbill/vb/api#SendEmail
'=========================================================================
Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
    
    '수신메일 주소
    receiverEmail = "test@test.com"
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증과 관련된 안내 SMS(단문) 문자를 재전송하는 함수로, 팝빌 사이트 [문자·팩스] > [문자] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/cashbill/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim sendNum As String
    Dim receiveNum As String
    Dim Contents As String
    
    '발신번호
    sendNum = "07043042991"
    
    '수신번호
    receiveNum = "010-111-222"
    
    ' 메시지 내용, 최대 90Byte (한글 45자), 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "알림 문자 내용, 최대 90Byte"
    
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, sendNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증을 팩스로 전송하는 함수로, 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/cashbill/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim sendNum As String
    Dim receiveNum As String
    
    '발신번호
    sendNum = "07043042991"
    
    '수신번호
    receiveNum = "010-111-222"
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, sendNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 관련 메일 항목에 대한 발송설정을 확인합니다.
' - https://docs.popbill.com/cashbill/vb/api#ListEmailConfig
'=========================================================================
Private Sub btnListemailconfig_Click()
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = CashbillService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "메일전송유형(EmailType) | 전송여부(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "CSH_ISSUE" Then
            tmp = tmp + "고객에게 현금영수증이 발행 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "CSH_CANCEL" Then
            tmp = tmp + "고객에게 현금영수증이 발행취소 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' 현금영수증 관련 메일 항목에 대한 발송설정을 수정합니다.
' - https://docs.popbill.com/cashbill/vb/api#UpdateEmailConfig
'
' 메일전송유형
' CSH_ISSUE : 고객에게 현금영수증이 발행 되었음을 알려주는 메일 입니다.
' CSH_CANCEL : 고객에게 현금영수증이 발행취소 되었음을 알려주는 메일 입니다.
'=========================================================================
Private Sub btnUpdateemailconfig_Click()
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '메일 전송 유형
    emailType = "CSH_ISSUE"

    '전송 여부 (True = 전송, False = 미전송)
    sendYN = True
    
    Set Response = CashbillService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 1건의 상세 정보 페이지의 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetPopUpURL
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPopUpURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 현금영수증 1건의 상세 정보 페이지(사이트 상단, 좌측 메뉴 및 버튼 제외)의 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetViewURL
'=========================================================================
Private Sub btnGetViewURl_Click()
    Dim URL As String
    
    URL = CashbillService.GetViewURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 현금영수증 1건을 인쇄하기 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetPrintURL
'=========================================================================
Private Sub btnGetPrintURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 1건의 현금영수증 인쇄 팝업 URL을 반환합니다. (공급받는자용)
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
'=========================================================================
Private Sub btnGetEPrintUrl_Click()
    Dim URL As String
    
    URL = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 다수건의 현금영수증을 인쇄하기 위한 페이지의 팝업 URL을 반환합니다. (최대 100건)
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetMassPrintURL
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim URL As String
    Dim KeyList As New Collection
    
    '문서번호 배열, 최대 100건
    KeyList.Add "20220101-01"
    KeyList.Add "20220101-02"
    KeyList.Add "20220101-03"
    KeyList.Add "20220101-04"
    
    URL = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList)
     
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 구매자가 수신하는 현금영수증 안내 메일의 하단에 버튼 URL 주소를 반환합니다.
' - 함수 호출로 반환 받은 URL에는 유효시간이 없습니다.
' - https://docs.popbill.com/cashbill/vb/api#GetMailURL
'=========================================================================
Private Sub btnGetMailURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 팝빌 현금영수증 임시문서함 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_TBOX_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 팝빌 현금영수증 발행문서함 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 팝빌 현금영수증 작성 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_WRITE_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

Private Sub Form_Load()
    CashbillService.Initialize linkID, SecretKey
    
    '연동환경설정값, True-개발용 False-상업용
    CashbillService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
    CashbillService.IPRestrictOnOff = True
    
    '로컬시스템 시간 사용여부 True-사용, False-미사용, 기본값(False)
    CashbillService.UseLocalTimeYN = False
End Sub

