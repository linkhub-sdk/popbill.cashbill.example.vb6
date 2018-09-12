VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 현금영수증 SDK 예제"
   ClientHeight    =   10860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   16170
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame7 
      Caption         =   "현금영수증 관련 기능"
      Height          =   7185
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   12330
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스 "
         Height          =   2295
         Left            =   1800
         TabIndex        =   46
         Top             =   1440
         Width           =   3135
         Begin VB.CommandButton btnDelete_ 
            Caption         =   "삭제"
            Height          =   375
            Left            =   1755
            Style           =   1  '그래픽
            TabIndex        =   49
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton btnCanceIssue_ 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   480
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
            Left            =   360
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
         Caption         =   " 문서 정보 "
         Height          =   2760
         Left            =   5880
         TabIndex        =   32
         Top             =   4125
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "공급받는자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "문서 내용 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "다량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "이메일(공급받는자) 링크 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   33
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   1890
         Left            =   9240
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
         Begin VB.CommandButton btnGetURL_SBOX 
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
         Caption         =   " 부가 서비스"
         Height          =   2775
         Left            =   2760
         TabIndex        =   24
         Top             =   4125
         Width           =   2895
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
            Caption         =   "문서 목록조회"
            Height          =   390
            Left            =   195
            TabIndex        =   50
            Top             =   2160
            Width           =   1845
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "문서 상세 정보"
            Height          =   390
            Left            =   195
            TabIndex        =   23
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "문서 이력"
            Height          =   390
            Left            =   195
            TabIndex        =   22
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "문서 정보(대량)"
            Height          =   390
            Left            =   210
            TabIndex        =   21
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "문서 정보"
            Height          =   390
            Left            =   210
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
         Begin VB.CommandButton btnRevokeRegistIssue_part 
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
         Caption         =   "관리번호 사용여부 확인"
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
         Caption         =   "관리번호( MgtKey) : "
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   15690
      Begin VB.Frame Frame15 
         Caption         =   "파트너과금 포인트"
         Height          =   1935
         Left            =   13080
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
         Height          =   1935
         Left            =   10920
         TabIndex        =   54
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnPopbillURL_CHRG 
            Caption         =   " 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련"
         Height          =   1935
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
         Begin VB.CommandButton btnListCorpInfo 
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
         Height          =   1935
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
         Height          =   1935
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
         Height          =   1935
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   2055
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
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL "
         Height          =   1935
         Left            =   6600
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
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
' - VB6 SDK 연동환경 설정방법 안내 :
' - 업데이트 일자 : 2017-08-30
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 27, 30번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================


Option Explicit


'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const LinkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'현금영수증 서비스 객체 생성
Private CashbillService As New PBCBService


'=========================================================================
' [발행완료] 상태의 현금영수증을 [발행취소] 합니다.
' - 발행취소는 국세청 전송전에만 가능합니다.
' - 발행취소된 형금영수증은 국세청에 전송되지 않습니다.
'=========================================================================

Private Sub btnCanceIssue__Click()
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
' [발행완료] 상태의 현금영수증을 [발행취소] 합니다.
' - 발행취소는 국세청 전송전에만 가능합니다.
' - 발행취소된 형금영수증은 국세청에 전송되지 않습니다.
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

Private Sub btnCancelIssue_rev_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
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
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 관리번호 중복여부를 확인합니다.
' - 관리번호는 1~24자리로 숫자, 영문 '-', '_' 조합으로 구성할 수 있습니다.
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
' 1건의 현금영수증을 삭제합니다.
' - 현금영수증을 삭제하면 사용된 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소]
'=========================================================================

Private Sub btnDelete__Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 현금영수증을 삭제합니다.
' - 현금영수증을 삭제하면 사용된 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소]
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
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
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
' 연동회원의 현금영수증 API 서비스 과금정보를 확인합니다.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = CashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 현금영수증 1건의 상세정보를 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[현금영수증 API 연동매뉴얼] > 4.1.
'   현금영수증 구성" 을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetDetailInfo_Click()

    Dim cbDetailInfo As PBCashbill
   
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "mgtKey (관리번호) : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "confirmNum (국세청승인번호) : " + cbDetailInfo.confirmNum + vbCrLf
    tmp = tmp + "tradeDate (거래일자) : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage (거래유형) : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType (현금영수증 형태) : " + cbDetailInfo.tradeType + vbCrLf
    tmp = tmp + "taxationType (과세형태) : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost (공급가액) : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax (세액) : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee (봉사료) : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount (거래금액) : " + cbDetailInfo.totalAmount + vbCrLf
    
    tmp = tmp + "franchiseCorpNum (발행자 사업자번호) : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseCorpName (발행자 상호) : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName (발행자 대표자명) : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr (발행자 주소) : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL (발행자 연락처) : " + cbDetailInfo.franchiseTEL + vbCrLf
    
    tmp = tmp + "identityNum (거래처 식별번호) : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName (고객명) : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName (상품명) : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber (주문번호) : " + cbDetailInfo.orderNumber + vbCrLf
    tmp = tmp + "email (고객 이메일) : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp (고객 휴대폰번호) : " + cbDetailInfo.hp + vbCrLf
    tmp = tmp + "smssendYN (알림문자 전송여부) : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
    
    tmp = tmp + "orgConfirmNum (원본현금영수증 국세청승인번호) : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (원본현금영수증 거래일자) : " + cbDetailInfo.orgTradeDate + vbCrLf
    tmp = tmp + "cancelType (취소사유) : " + CStr(cbDetailInfo.cancelType) + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 현금영수증 인쇄(공급받는자) URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 현금영수증 상태/요약 정보를 확인합니다.
' - 응답항목에 대한 자세한 정보는 "[현금영수증 API 연동매뉴얼] > 4.2.
'   현금영수증 상태정보 구성"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
 
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "itemKey (아이템키) : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey (문서관리번호) : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate (거래일자) : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "issueDT (발행일시) : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "regDT (등록일시) : " + cbInfo.regDT + vbCrLf
    tmp = tmp + "taxationType (과세형태) : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount (거래금액) : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "tradeUsage (거래용도) : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType (현금영수증 형태) : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "stateCode (상태코드) : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (상태변경일시) : " + cbInfo.stateDT + vbCrLf
    
    tmp = tmp + "identityNum (거래처 식별번호) : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "itemName (상품명) : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "customerName (고객명) : " + cbInfo.customerName + vbCrLf
    
    tmp = tmp + "confirmNum (국세청승인번호) : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "ntssendDT (국세청 전송일시) : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (국세청 처리결과 수신일시) : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode (국세청 처리결과 상태코드) : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage (국세청 처리결과 메시지) : " + cbInfo.ntsresultMessage + vbCrLf
    tmp = tmp + "orgConfirmNum (원본 현금영수증 국세청 승인번호) : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (원본 현금영수증 거래일자) : " + cbInfo.orgTradeDate + vbCrLf
    
    tmp = tmp + "printYN (인쇄여부) : " + CStr(cbInfo.printYN) + vbCrLf
   
    MsgBox tmp
    
    
End Sub

'=========================================================================
' 다수건의 현금영수증 상태/요약 정보를 확인합니다.
' - 응답항목에 대한 자세한 정보는 "[현금영수증 API 연동매뉴얼] > 4.2.
'   현금영수증 상태정보 구성"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    
    '현금영수증 관리번호 배열, 최대 1000건
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT" + vbCrLf
    
    Dim info As PBCbInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxationType + " | " + info.tradeDate + " | " + info.confirmNum + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 현금영수증 상태 변경이력을 확인합니다.
' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
'   "[현금영수증 API 연동매뉴얼] > 3.4.4 상태 변경이력 확인"
'   을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBCbLog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 공급받는자 메일링크 URL을 반환합니다.
' - 메일링크 URL은 유효시간이 존재하지 않습니다.
'=========================================================================

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 다수건의 현금영수증 인쇄팝업 URL을 반환합니다.
' 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    url = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 파트너 포인트충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = CashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 현금영수증 보기 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetPopUpURL_Click()
    Dim url As String
    
    url = CashbillService.GetPopUpURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 현금영수증 인쇄팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 현금영수증 > 발행문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 현금영수증 > 임시(연동)문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 현금영수증 > 현금영수증 작성 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 임시저장 현금영수증을 발행처리합니다.
' - 발행일 기준 오후 5시 이전에 발행된 현금영수증은 다음날 오후 2시에 국세청
'   전송결과를 확인할 수 있습니다.
' - 현금영수증 국세청 전송 정책에 대한 정보는 "[현금영수증 API 연동매뉴얼]
'   > 1.4. 국세청 전송정책"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "발행메모"
    
    Set Response = CashbillService.Issue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub


Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "6748500389"
    
    '대표자성명, 최대 30자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 70자
    joinData.corpName = "회원상호"
    
    '주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 40자
    joinData.bizType = "업태"
    
    '종목, 최대 40자
    joinData.bizClass = "종목"
    
    '아이디, 6자이상 20자 미만
    joinData.id = "testkorea_1011"
    
    '비밀번호, 6자이상 20자 미만
    joinData.pwd = "pwd_must_be_long_enough"
    
    '담당자명, 최대 30자
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    '담당자 메일, 최대 70자
    joinData.ContactEmail = "test@test.com"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = CashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================

Private Sub btnListCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = CashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname(대표자성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(회사명) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub
'=========================================================================
' 현금영수증 관련 메일전송 항목에 대한 전송여부를 목록으로 반환합니다
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
' 현금영수증 관련 메일전송 항목에 대한 전송여부를 수정합니다.
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
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 20자 미만
    joinData.id = "testkorea_20161010"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.pwd = "test@test.com"
    
    '담당자명, 최대 30자
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
    
    '담당자 팩스번호
    joinData.fax = "070-1234-1234"
    
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
        
    Set Response = CashbillService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 현금영수증을 임시저장 합니다.
' - [임시저장] 상태의 현금영수증은 발행(Issue API)을 호출해야만 국세청에
'   전송됩니다.
' - 발행일 기준 오후 5시 이전에 발행된 현금영수증은 다음날 오후 2시에 국세청
'   전송결과를 확인할 수 있습니다.
' - 현금영수증 국세청 전송 정책에 대한 정보는 "[현금영수증 API 연동매뉴얼]
'   > 1.4. 국세청 전송정책"을 참조하시기 바랍니다.
' - 취소현금영수증 작성방법 안내 - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRegister_Click()
    Dim Cashbill As New PBCashbill
    
    '현금영수증 관리번호, 1~24자리 영문,숫자조합으로 사업자별로 중복되지 않도록 구성
    Cashbill.mgtKey = txtMgtKey.Text
    
    '현금영수증 형태, [승인거래, 취소거래] 중 기재
    Cashbill.tradeType = "승인거래"
    
    '[취소거래시 필수] 원본 국세청승인번호
    '문서정보(GetInfo API)의 응답항목중 국세청승인번호(confirmNum)를 확인하여 기재
    Cashbill.orgConfirmNum = ""
    
    '[취소거래시 필수] 원본 현금영수증 거래일자
    '문서정보(GetInfo API)의 응답항목중 거래일자(tradeDate)를 확인하여 기재
    Cashbill.orgTradeDate = ""
    
    '발행자 사업자번호, "-" 제외 10자리
    Cashbill.franchiseCorpNum = "1234567890"
    
    '발행자 상호명
    Cashbill.franchiseCorpName = "발행자 상호"
    
    '발행자 대표자 성명
    Cashbill.franchiseCEOName = "발행자 대표자"
    
    '발행자 주소
    Cashbill.franchiseAddr = "발행자 주소"
    
    '발행자 연락처
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '거래유형, [소득공제용, 지출증빙용] 중 기재
    Cashbill.tradeUsage = "소득공제용"
    
    '거래처 식별번호, 거래유형에 따라 작성
    '소득공제용 - 주민등록/휴대폰/카드번호 기재가능
    '지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호 기재가능
    Cashbill.identityNum = "0101112222"
    
    '과세형태, [과세, 비과세] 중 기재
    Cashbill.taxationType = "과세"
    
    '공급가액
    Cashbill.supplyCost = "10000"
    
    '봉사료
    Cashbill.serviceFee = "0"
    
    '세액
    Cashbill.tax = "1000"
    
    '합계금액, 공급가액 + 봉사료 + 세액
    Cashbill.totalAmount = "11000"
    
    '주문고객명
    Cashbill.customerName = "고객명"
    
    '상품명
    Cashbill.itemName = "상품명"
    
    '주문번호
    Cashbill.orderNumber = "주문번호"
    
    '고객이메일
    Cashbill.email = "test@test.com"
    
    '고객휴대폰번호
    Cashbill.hp = "010-111-222"
    
    '현금영수증 발행 알림문자 전송여부
    Cashbill.smssendYN = False
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.Register(txtCorpNum.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    

End Sub

'=========================================================================
' 1건의 현금영수증을 즉시발행합니다.
' - 발행일 기준 오후 5시 이전에 발행된 현금영수증은 다음날 오후 2시에 국세청
'   전송결과를 확인할 수 있습니다.
' - 현금영수증 국세청 전송 정책에 대한 정보는 "[현금영수증 API 연동매뉴얼]
'   > 1.4. 국세청 전송정책"을 참조하시기 바랍니다.
' - 취소현금영수증 작성방법 안내 - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Cashbill As New PBCashbill
    
    '현금영수증 관리번호, 1~24자리 영문,숫자조합으로 사업자별로 중복되지 않도록 구성
    Cashbill.mgtKey = txtMgtKey.Text
    
    '현금영수증 형태, [승인거래, 취소거래] 중 기재
    Cashbill.tradeType = "승인거래"
    
    '[취소거래시 필수] 원본 국세청승인번호
    '문서정보(GetInfo API)의 응답항목중 국세청승인번호(confirmNum)를 확인하여 기재
    Cashbill.orgConfirmNum = ""
    
    '[취소거래시 필수] 원본 현금영수증 거래일자
    '문서정보(GetInfo API)의 응답항목중 거래일자(tradeDate)를 확인하여 기재
    Cashbill.orgTradeDate = ""
    
    '발행자 사업자번호, "-" 제외 10자리
    Cashbill.franchiseCorpNum = txtCorpNum.Text
    
    '발행자 상호명
    Cashbill.franchiseCorpName = "발행자 상호"
    
    '발행자 대표자 성명
    Cashbill.franchiseCEOName = "발행자 대표자"
    
    '발행자 주소
    Cashbill.franchiseAddr = "발행자 주소"
    
    '발행자 연락처
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '거래유형, [소득공제용, 지출증빙용] 중 기재
    Cashbill.tradeUsage = "소득공제용"
    
    '거래처 식별번호, 거래유형에 따라 작성
    '소득공제용 - 주민등록/휴대폰/카드번호 기재가능
    '지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호 기재가능
    Cashbill.identityNum = "0101112222"
    
    '과세형태, [과세, 비과세] 중 기재
    Cashbill.taxationType = "과세"
    
    '공급가액
    Cashbill.supplyCost = "10000"
    
    '봉사료
    Cashbill.serviceFee = "0"
    
    '세액
    Cashbill.tax = "1000"
    
    '합계금액, 공급가액 + 봉사료 + 세액
    Cashbill.totalAmount = "11000"
    
    '주문고객명
    Cashbill.customerName = "고객명"
    
    '상품명
    Cashbill.itemName = "상품명"
    
    '주문번호
    Cashbill.orderNumber = "주문번호"
    
    '고객이메일
    Cashbill.email = "test@test.com"
    
    '고객휴대폰번호
    Cashbill.hp = "010-111-222"
    
    '현금영수증 발행 알림문자 전송여부
    Cashbill.smssendYN = False
        
    Dim Response As PBResponse
    
    Set Response = CashbillService.RegistIssue(txtCorpNum.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 1건의 취소현금영수증을 즉시발행합니다.
' - 발행일 기준 오후 5시 이전에 발행된 현금영수증은 다음날 오후 2시에 국세청
'   전송결과를 확인할 수 있습니다.
' - 현금영수증 국세청 전송 정책에 대한 정보는 "[현금영수증 API 연동매뉴얼]
'   > 1.4. 국세청 전송정책"을 참조하시기 바랍니다.
' - 취소현금영수증 작성방법 안내 - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRevokeRegistIssue_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    
    
    '원본현금영수증 승인번호
    orgConfirmNum = "820116333"
    
    '원본현금영수증 거래일자
    orgTradeDate = "20170711"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 (부분 취소현금영수증을 즉시발행합니다.
' - 발행일 기준 오후 5시 이전에 발행된 현금영수증은 다음날 오후 2시에 국세청
'   전송결과를 확인할 수 있습니다.
' - 현금영수증 국세청 전송 정책에 대한 정보는 "[현금영수증 API 연동매뉴얼]
'   > 1.4. 국세청 전송정책"을 참조하시기 바랍니다.
' - 취소현금영수증 작성방법 안내 - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRevokeRegistIssue_part_Click()
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
    orgConfirmNum = "820116333"
    
    '원본현금영수증 거래일자
    orgTradeDate = "20170711"
    
    '안내문자 전송여부
    smssendYN = False
    
    '메모
    memo = "즉시발행 메모"
    
    '부분취소여부, True-부분취소, False-전체취소
    isPartCancel = True
    
    '취소사유(Integer), 1-거래취소, 2-오류발급취소, 3-기타
    cancelType = 1
    
    '[취소] 공급가액
    supplyCost = "3000"
    
    '[취소] 세액
    tax = "300"
    
    '[취소] 봉사료
    serviceFee = "0"
    
    '[취소] 합계금액
    totalAmount = "3300"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 검색조건을 사용하여 현금영수증 목록을 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[현금영수증 API 연동매뉴얼] >
'   4.2. 현금영수증 상태정보 구성" 을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnSearch_Click()
    Dim cbSearchList As PBCBSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim taxationType As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim QString As String
    Dim Order As String
    
    '[필수] 일자유형, R-등록일자, T-거래일자 I-발행일자
    DType = "T"
    
    '[필수] 시작일자, 형식(yyyyMMdd)
    SDate = "20160901"
    
    '[필수] 종료일자, 형식(yyyyMMdd)
    EDate = "20161031"
    
    '전송상태코드 배열, 미기재시 전체조회, 2,3번째 자리 와일드카드(*) 가능
    '[참조] 현금영수증 API 연동매뉴열 "5.1. 현금영수증 상태코드"
    state.Add "2**"
    state.Add "3**"
    state.Add "4**"
    
    '현금영수증 형태 배열, N-일반 현금영수증, C-취소 현금영수증
    tradeType.Add "N"
    tradeType.Add "C"
    
    '거래유형 배열, P-소득공제, C-제출증빙
    tradeUsage.Add "P"
    tradeUsage.Add "C"
    
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
    
    Set cbSearchList = CashbillService.Search(txtCorpNum.Text, DType, SDate, EDate, state, tradeType, _
                                tradeUsage, taxationType, Page, PerPage, Order, QString)
     
    If cbSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (응답코드)  : " + CStr(cbSearchList.code) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + cbSearchList.message + vbCrLf + vbCrLf + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(cbSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(cbSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(cbSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(cbSearchList.pageCount) + vbCrLf
    
    tmp = tmp + "ItemKey | MgtKey | TradeDate | TradeUsage | IssueDT | CustomerName | ItemName | IdentityNum | TaxationType | TotalAmount | tradeType | StateCode | TxationType | TradeDate | confirmNum " + vbCrLf
    
    Dim info As PBCbInfo
    
    For Each info In cbSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + info.mgtKey + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.tradeUsage + " | "
        tmp = tmp + info.issueDT + " | "
        tmp = tmp + info.customerName + " | "
        tmp = tmp + info.itemName + " | "
        tmp = tmp + info.identityNum + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | "
        tmp = tmp + info.tradeType + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.confirmNum + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 발행 안내메일을 재전송합니다.
'=========================================================================

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiveEmail As String
    
    '수신메일주소
    receiveEmail = "test@test.com"
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, receiveEmail)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증을 팩스전송합니다.
' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [팩스] > [전송내역]
'   메뉴에서 전송결과를 확인할 수 있습니다.
'=========================================================================

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiveNum As String
    
    '발신번호
    senderNum = "07043042991"
    
    '수신번호
    receiveNum = "010-111-222"
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, senderNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 알림문자를 전송합니다. (단문/SMS- 한글 최대 45자)
' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [전송내역] 탭에서
'   전송결과를 확인할 수 있습니다.
'=========================================================================

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiveNum As String
    Dim Contents As String
    
    '발신번호
    senderNum = "07075103710"
    
    '수신번호
    receiveNum = "010-111-222"
    
    '문자메시지 내용, 90Byte를 초과한 내용은 삭제되어 전송됨
    Contents = "알림 문자 내용, 최대 90Byte"
      
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, senderNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 현금영수증 발행단가를 확인합니다.
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
' 1건의 현금영수증을 수정합니다.
' - [임시저장] 상태의 현금영수증만 수정할 수 있습니다.
' - 국세청에 신고된 현금영수증은 수정할 수 없으며, 취소 현금영수증을 발행하여
'   취소처리 할 수 있습니다.
' - 취소현금영수증 작성방법 안내 - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnUpdate_Click()
    Dim Cashbill As New PBCashbill
    Dim Response As PBResponse
    
    '현금영수증 관리번호, 1~24자리 영문,숫자조합으로 사업자별로 중복되지 않도록 구성
    Cashbill.mgtKey = txtMgtKey.Text
    
    '현금영수증 형태, [승인거래, 취소거래] 중 기재
    Cashbill.tradeType = "승인거래"
    
    '발행자 사업자번호, "-" 제외 10자리
    Cashbill.franchiseCorpNum = "1234567890"
    
    '발행자 상호명
    Cashbill.franchiseCorpName = "발행자 상호_수정"
    
    '발행자 대표자 성명
    Cashbill.franchiseCEOName = "발행자 대표자_수정"
    
    '발행자 주소
    Cashbill.franchiseAddr = "발행자 주소"
    
    '발행자 연락처
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '거래유형, [소득공제용, 지출증빙용] 중 기재
    Cashbill.tradeUsage = "소득공제용"
    
    '거래처 식별번호, 거래유형에 따라 작성
    '소득공제용 - 주민등록/휴대폰/카드번호 기재가능
    '지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호 기재가능
    Cashbill.identityNum = "01041680206"
    
    '과세형태, [과세, 비과세] 중 기재
    Cashbill.taxationType = "과세"
    
    '공급가액
    Cashbill.supplyCost = "10000"
    
    '봉사료
    Cashbill.serviceFee = "0"
    
    '세액
    Cashbill.tax = "1000"
    
    '합계금액, 공급가액 + 봉사료 + 세액
    Cashbill.totalAmount = "11000"
    
    '주문고객명
    Cashbill.customerName = "고객명"
    
    '상품명
    Cashbill.itemName = "상품명"
    
    '주문번호
    Cashbill.orderNumber = "주문번호"
    
    '고객이메일
    Cashbill.email = "test@test.com"
    
    '고객휴대폰번호
    Cashbill.hp = "010-111-222"
    
    '현금영수증 발행 알림문자 전송여부
    Cashbill.smssendYN = False
    
    Set Response = CashbillService.Update(txtCorpNum.Text, txtMgtKey.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-1234-1234"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, Ture-회사조회, False-개인조
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False

                
    Set Response = CashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자성명, 최대 30자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 70자
    CorpInfo.corpName = "상호"
    
    ' 주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 40자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 40자
    CorpInfo.bizClass = "종목"
    
    Set Response = CashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(CashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub




Private Sub Form_Load()
    CashbillService.Initialize LinkID, SecretKey
    
    '연동환경 설정값 True-테스트용, False-상업용
    CashbillService.IsTest = True
End Sub

