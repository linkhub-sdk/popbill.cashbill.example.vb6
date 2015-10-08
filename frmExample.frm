VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 현금영수증 SDK 예제"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame7 
      Caption         =   "현금영수증 관련 기능"
      Height          =   6705
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   10650
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스 "
         Height          =   2295
         Left            =   1320
         TabIndex        =   53
         Top             =   840
         Width           =   3135
         Begin VB.CommandButton btnDelete_ 
            Caption         =   "삭제"
            Height          =   375
            Left            =   1755
            Style           =   1  '그래픽
            TabIndex        =   56
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton btnCanceIssue_ 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   480
            Style           =   1  '그래픽
            TabIndex        =   55
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   430
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   54
            Top             =   460
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
         Left            =   4800
         TabIndex        =   36
         Top             =   3525
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "공급받는자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   41
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "문서 내용 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   40
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   39
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "다량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   38
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "이메일(공급받는자) 링크 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   1890
         Left            =   8160
         TabIndex        =   32
         Top             =   3525
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "발행 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "매출 작성"
            Height          =   390
            Left            =   210
            TabIndex        =   33
            Top             =   1275
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " 부가 서비스"
         Height          =   1935
         Left            =   2760
         TabIndex        =   28
         Top             =   3525
         Width           =   1815
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   31
            Top             =   390
            Width           =   1335
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   30
            Top             =   825
            Width           =   1335
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   390
            Left            =   210
            TabIndex        =   29
            Top             =   1290
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " 문서 정보 "
         Height          =   2295
         Left            =   240
         TabIndex        =   23
         Top             =   3525
         Width           =   2265
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "문서 상세 정보"
            Height          =   390
            Left            =   195
            TabIndex        =   27
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "문서 이력"
            Height          =   390
            Left            =   195
            TabIndex        =   26
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "문서 정보(대량)"
            Height          =   390
            Left            =   210
            TabIndex        =   25
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "문서 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   24
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "임시저장-발행 프로세스 "
         Height          =   2385
         Left            =   4800
         TabIndex        =   18
         Top             =   840
         Width           =   3735
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   525
            Left            =   1230
            Style           =   1  '그래픽
            TabIndex        =   43
            Top             =   1140
            Width           =   1020
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   1215
            Style           =   1  '그래픽
            TabIndex        =   42
            Top             =   1830
            Width           =   960
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "등록"
            Height          =   375
            Left            =   1305
            Style           =   1  '그래픽
            TabIndex        =   21
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "수정"
            Height          =   375
            Left            =   2265
            Style           =   1  '그래픽
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "삭제"
            Height          =   375
            Left            =   2490
            Style           =   1  '그래픽
            TabIndex        =   19
            Top             =   1830
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   1740
            X2              =   1740
            Y1              =   2100
            Y2              =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   465
            TabIndex        =   22
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   3135
         End
         Begin VB.Line Line3 
            X1              =   2925
            X2              =   2925
            Y1              =   2010
            Y2              =   615
         End
         Begin VB.Line Line2 
            X1              =   1710
            X2              =   3120
            Y1              =   2010
            Y2              =   2010
         End
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "관리번호 사용여부 확인"
         Height          =   375
         Left            =   6150
         TabIndex        =   17
         Top             =   255
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   3300
         TabIndex        =   16
         Top             =   285
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "관리번호( MgtKey) : "
         Height          =   180
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   11010
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련"
         Height          =   2295
         Left            =   8880
         TabIndex        =   50
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnListCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   2295
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   1560
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
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
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
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   1560
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
         Begin VB.CommandButton btnPopbillURL_CHRG 
            Caption         =   " 포인트 충전 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   495
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
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'링크아이디
Private Const linkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private CashbillService As New PBCBService

Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub


Private Sub btnCanceIssue__Click()
    Dim Response As PBResponse
  
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
  
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIssue_rev_Click()
    Dim Response As PBResponse
       
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub



Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.CheckMgtKeyInUse(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnDelete__Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnDelete_Click()
    Dim Response As PBResponse
  
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
    
End Sub

Private Sub btnGetDetailInfo_Click()

    Dim cbDetailInfo As PBCashbill
   
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "mgtKey : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType : " + cbDetailInfo.tradeType + vbCrLf
    
    tmp = tmp + "taxationType : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount : " + cbDetailInfo.totalAmount + vbCrLf
    
    tmp = tmp + "franchiseCorpNum : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseCorpName : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL : " + cbDetailInfo.franchiseTEL + vbCrLf
    
    tmp = tmp + "identityNum : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber : " + cbDetailInfo.orderNumber + vbCrLf
    
    tmp = tmp + "email : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp : " + cbDetailInfo.hp + vbCrLf
    tmp = tmp + "fax : " + cbDetailInfo.fax + vbCrLf
    tmp = tmp + "smssendYN : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
    tmp = tmp + "faxsendYN : " + CStr(cbDetailInfo.faxsendYN) + vbCrLf
    
    tmp = tmp + "confirmNum : " + cbDetailInfo.confirmNum + vbCrLf
    
    tmp = tmp + "orgConfirmNum : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate : " + cbDetailInfo.orgTradeDate + vbCrLf
    
    
    
    
    MsgBox tmp
    
End Sub


Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub


Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
 
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
   
    tmp = tmp + "itemKey : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "issueDT : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "customerName : " + cbInfo.customerName + vbCrLf
    tmp = tmp + "itemName : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "identityNum : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "taxationType : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "tradeUsage : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "stateCode : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT : " + cbInfo.stateDT + vbCrLf
    tmp = tmp + "printYN : " + CStr(cbInfo.printYN) + vbCrLf
    tmp = tmp + "confirmNum : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "orgTradeDate : " + cbInfo.orgTradeDate + vbCrLf
    tmp = tmp + "orgConfirmNum : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "ntssendDT : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresult : " + cbInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsresultDT : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage : " + cbInfo.ntsresultMessage + vbCrLf
   
    tmp = tmp + "regDT : " + cbInfo.regDT + vbCrLf
    
    MsgBox tmp
    
    
End Sub

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
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

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcCorpName | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBCbLog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    url = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList, txtUserID.Text)
     
    If url = "" Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub


Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnIssue_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Issue(txtCorpNum.Text, txtMgtKey.Text, "발행메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub


Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID '링크 아이디
    joinData.CorpNum = "1231212312" '사업자번호 "-" 제외.
    joinData.ceoname = "대표자성명"
    joinData.corpName = "회원상호"
    joinData.addr = "주소"
    joinData.ZipCode = "500-100"
    joinData.bizType = "업태"
    joinData.bizClass = "업종"
    joinData.id = "userid"      '6자 이상 20자 미만.
    joinData.pwd = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub


Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = CashbillService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnListCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = CashbillService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.id = "testkorea_20151007"      '담당자 아이디
    joinData.pwd = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = CashbillService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnRegister_Click()
    Dim Cashbill As New PBCashbill
    
    Cashbill.mgtKey = txtMgtKey.Text        '발행자별 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
    Cashbill.tradeType = "승인거래"         '승인거래 or 취소거래
    Cashbill.franchiseCorpNum = "1231212312"
    Cashbill.franchiseCorpName = "발행자 상호"
    Cashbill.franchiseCEOName = "발행자 대표자"
    Cashbill.franchiseAddr = "발행자 주소"
    Cashbill.franchiseTEL = "070-1234-1234"
    Cashbill.identityNum = "01041680206"
    Cashbill.customerName = "고객명"
    Cashbill.itemName = "상품명"
    Cashbill.orderNumber = "주문번호"
    Cashbill.email = "test@test.com"
    Cashbill.hp = "111-1234-1234"
    Cashbill.fax = "777-444-3333"
    Cashbill.serviceFee = "0"
    Cashbill.supplyCost = "10000"
    Cashbill.tax = "1000"
    Cashbill.totalAmount = "11000"
    Cashbill.tradeUsage = "소득공제용"      '소득공제용 or 지출증빙용
    Cashbill.taxationType = "과세"          '과세 or 비과세
    
    Cashbill.smssendYN = False
    Cashbill.faxsendYN = False
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.Register(txtCorpNum.Text, Cashbill, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    

End Sub


Private Sub btnRegistIssue_Click()
    Dim Cashbill As New PBCashbill
    
        
    Cashbill.mgtKey = txtMgtKey.Text        '발행자별 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
    Cashbill.tradeType = "승인거래"         '승인거래 or 취소거래
    Cashbill.franchiseCorpNum = "1234567890"
    Cashbill.franchiseCorpName = "발행자 상호"
    Cashbill.franchiseCEOName = "발행자 대표자"
    Cashbill.franchiseAddr = "발행자 주소"
    Cashbill.franchiseTEL = "070-1234-1234"
    Cashbill.identityNum = "01043245117"
    Cashbill.customerName = "고객명"
    Cashbill.itemName = "상품명"
    Cashbill.orderNumber = "주문번호"
    Cashbill.email = "test@test.com"
    Cashbill.hp = "111-1234-1234"
    Cashbill.fax = "777-444-3333"
    Cashbill.serviceFee = "0"
    Cashbill.supplyCost = "10000"
    Cashbill.tax = "1000"
    Cashbill.totalAmount = "11000"
    Cashbill.tradeUsage = "소득공제용"      '소득공제용 or 지출증빙용
    Cashbill.taxationType = "과세"          '과세 or 비과세
    Cashbill.memo = "즉시발행 메모"
    
    Cashbill.smssendYN = False
    Cashbill.faxsendYN = False
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.RegistIssue(txtCorpNum.Text, Cashbill, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, "test@test.com", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, "07075106766", "111-2222-4444", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, "07075106766", "111-2222-4444", "문자 내용 최대 90Byte", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = CashbillService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "발행단가 : " + CStr(unitCost)
End Sub

Private Sub btnUpdate_Click()

    Dim Cashbill As New PBCashbill
    Dim Response As PBResponse
    
    Set Response = CashbillService.Update(txtCorpNum.Text, txtMgtKey.Text, Cashbill, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
    Set Response = CashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.ceoname = "대표자"         '대표자명
    CorpInfo.corpName = "상호_수정"          '상호명
    CorpInfo.addr = "서울특별시"        '주소
    CorpInfo.bizType = "업태"           '업태
    CorpInfo.bizClass = "업종"          '업종
    
    Set Response = CashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub Form_Load()
    CashbillService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True-테스트용, False-상업용
    CashbillService.IsTest = True
End Sub

