VERSION 5.00
Begin VB.Form PopbillEasyFinBankExample 
   Caption         =   "팝빌 간편 계좌조회 API Example"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   17865
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame6 
      Caption         =   "계좌조회 관련 API"
      Height          =   6735
      Left            =   240
      TabIndex        =   26
      Top             =   3360
      Width           =   17175
      Begin VB.ListBox searchInfo 
         Height          =   2940
         Left            =   360
         TabIndex        =   44
         Top             =   3360
         Width           =   15135
      End
      Begin VB.TextBox txtJobID 
         Height          =   375
         Left            =   1440
         TabIndex        =   42
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Frame Frame10 
         Caption         =   "정액제 관리"
         Height          =   2295
         Left            =   12360
         TabIndex        =   30
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "정액제 서비스 상태 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "정액제 서비스 신청 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "계좌관리"
         Height          =   2295
         Left            =   5160
         TabIndex        =   29
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton btnRevokeCloseBankAccount 
            Caption         =   "정액제 해지신청 취소"
            Height          =   495
            Left            =   4440
            TabIndex        =   49
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnCloseBankAccount 
            Caption         =   "계좌 정액제 해지신청"
            Height          =   495
            Left            =   4440
            TabIndex        =   48
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBankAccountInfo 
            Caption         =   "계좌정보 확인"
            Height          =   495
            Left            =   2280
            TabIndex        =   47
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnUpdateBankAccount 
            Caption         =   "계좌정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton btnRegistBankAccount 
            Caption         =   "계좌 등록"
            Height          =   495
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnListBankAccount 
            Caption         =   "계좌 목록 확인"
            Height          =   495
            Left            =   2280
            TabIndex        =   38
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBankAccountMgtURL 
            Caption         =   "계좌관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   2055
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "거래내역 관리"
         Height          =   2295
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "거래내역 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "거래내역 요약정보 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnSaveMemo 
            Caption         =   "거래내역 메모저장"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   1560
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "거래내역 수집"
         Height          =   2295
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "수집 상태 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "수집 상태 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "수집 요청"
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(작업아이디는 '수집 요청' 호출시 반환됩니다 )"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "작업아이디 : "
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "회사정보 수정"
      Height          =   410
      Left            =   14880
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame15 
      Caption         =   "회사정보 관련"
      Height          =   1935
      Left            =   14760
      TabIndex        =   6
      Top             =   1080
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "회사정보 조회"
         Height          =   410
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "담당자 목록 조회"
      Height          =   410
      Left            =   9960
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "담당자 정보 수정"
      Height          =   410
      Left            =   9960
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID 중복 확인"
      Height          =   410
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2535
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   17175
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   1935
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   1935
         Left            =   9600
         TabIndex        =   17
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   1935
         Left            =   12000
         TabIndex        =   15
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "연동과금 포인트"
         Height          =   1935
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "파트너과금 포인트"
         Height          =   1935
         Left            =   6960
         TabIndex        =   9
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2295
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Text            =   "testkorea"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   255
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4560
      TabIndex        =   25
      Top             =   315
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 :"
      Height          =   180
      Left            =   360
      TabIndex        =   24
      Top             =   315
      Width           =   1920
   End
End
Attribute VB_Name = "PopbillEasyFinBankExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' 팝빌 계좌조회 API VB 6.0 SDK Example
'
' - 업데이트 일자 : 2020-07-20
' - 연동 기술지원 연락처 : 1600-8536 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
'=========================================================================

Option Explicit

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'계좌조회 서비스 클래스 선언
Private easyFinBankService As New PBEasyFinBankService

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = easyFinBankService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/easyfinbank/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = easyFinBankService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 계좌의 정액제 해지를 요청합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#CloseBankAccount
'=========================================================================
Private Sub btnCloseBankAccount_Click()
    
    Dim Response As PBResponse
    Dim BankCode As String
    Dim AccountNumber As String
    Dim CloseType As String
    
    ' [필수] 은행코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    BankCode = ""
    
    ' [필수] 계좌번호 하이픈('-') 제외
    AccountNumber = ""

    ' 해지유형, "일반", "중도" 중 선택 기재
    ' 일반해지 - 이용중인 정액제 사용기간까지 이용후 정지
    ' 중도해지 - 요청일 기준으로 정지, 정액제 잔여기간은 일할로 계산되어 포인트 환불 (무료 이용기간 중 중도해지 시 전액 환불)
    CloseType = "중도"
    
    Set Response = easyFinBankService.CloseBankAccount(txtCorpNum.Text, BankCode, AccountNumber, CloseType)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = easyFinBankService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "연동회원 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 팝빌에 등록된 계좌 정보를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBankAccountInfo
'=========================================================================
Private Sub btnGetBankAccountInfo_Click()
    Dim AccountInfo As PBEasyFinBankAccount
    Dim tmp As String
    Dim BankCode As String
    Dim AccountNumber As String
    
    ' [필수] 은행코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    BankCode = ""
    
    ' [필수] 계좌번호 하이픈('-') 제외
    AccountNumber = ""
    
    Set AccountInfo = easyFinBankService.GetBankAccountInfo(txtCorpNum.Text, BankCode, AccountNumber)
     
    If AccountInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "bankCode (은행코드) : " + AccountInfo.BankCode + vbCrLf
    tmp = tmp + "accountNumber (계좌번호) : " + AccountInfo.AccountNumber + vbCrLf
    tmp = tmp + "accountName (계좌별칭) : " + AccountInfo.AccountName + vbCrLf
    tmp = tmp + "accountType (계좌유형) : " + AccountInfo.AccountType + vbCrLf
    tmp = tmp + "state (정액제 상태) : " + CStr(AccountInfo.state) + vbCrLf
    tmp = tmp + "regDT (등록일시) : " + AccountInfo.regDT + vbCrLf
    tmp = tmp + "contractDT (정액제 서비스 시작일시) : " + AccountInfo.contractDT + vbCrLf
    tmp = tmp + "baseDate (자동연장 결제일) : " + CStr(AccountInfo.baseDate) + vbCrLf
    tmp = tmp + "useEndDate (정액제 서비스 종료일자) : " + AccountInfo.useEndDate + vbCrLf
    tmp = tmp + "contractState (정액제 서비스 상태) : " + CStr(AccountInfo.contractState) + vbCrLf
    tmp = tmp + "closeRequestYN (정액제 해지신청 여부) : " + CStr(AccountInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (정액제 사용제한 여부) : " + CStr(AccountInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (정액제 만료시 해지여부) : " + CStr(AccountInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaiedYN (미수금 보유 여부) : " + CStr(AccountInfo.unPaidYN) + vbCrLf
    tmp = tmp + "memo (메모) : " + AccountInfo.Memo
    
    MsgBox tmp
End Sub

'=========================================================================
' 계좌 등록, 수정 및 삭제할 수 있는 계좌 관리 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBankAccountMgtURL
'=========================================================================
Private Sub btnGetBankAccountMgtURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetBankAccountMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 계좌조회 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = easyFinBankService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (월정액단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = easyFinBankService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 계좌조회 정액제 서비스 신청 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetFlatRatePopUpURL
'=========================================================================
Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetFlatRatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 계좌조회 정액제 서비스 상태를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetFlatRateState
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBEasyFinBankFlatRate
    Dim tmp As String
    Dim BankCode As String
    Dim AccountNumber As String
    
    '은행코드
    BankCode = "0048"
    
    '팝빌에 등록된 계좌번호
    AccountNumber = "131020538645"
    
    Set flatRateInfo = easyFinBankService.GetFlatRateState(txtCorpNum.Text, BankCode, AccountNumber)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "referencdeID (계좌아이디) : " + flatRateInfo.referenceID + vbCrLf
    tmp = tmp + "contractDT (정액제 서비스 시작일시) : " + flatRateInfo.contractDT + vbCrLf
    tmp = tmp + "useEndDate (정액제 서비스 종료일) : " + flatRateInfo.useEndDate + vbCrLf
    tmp = tmp + "baseDate (자동연장 결제일) : " + CStr(flatRateInfo.baseDate) + vbCrLf
    tmp = tmp + "state (정액제 서비스 상태) : " + CStr(flatRateInfo.state) + vbCrLf
    tmp = tmp + "closeRequestYN (서비스 해지신청 여부) : " + CStr(flatRateInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (서비스 사용제한 여부) : " + CStr(flatRateInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (서비스만료시 해지여부 ) : " + CStr(flatRateInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaidYN (미수금 보유 여부) : " + CStr(flatRateInfo.unPaidYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' RequestJob(수집 요청)를 통해 반환 받은 작업아이디의 상태를 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetJobState
'=========================================================================
Private Sub btnGetJobState_Click()
    Dim jobInfo As PBEasyFinBankJobState
    Dim tmp As String
    
    Set jobInfo = easyFinBankService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "jobID(작업아이디) : " + jobInfo.jobID + vbCrLf
    tmp = tmp + "jobState(수집상태) : " + CStr(jobInfo.jobState) + vbCrLf
    tmp = tmp + "startDate(시작일자) : " + jobInfo.startDate + vbCrLf
    tmp = tmp + "endDate(종료일자) : " + jobInfo.endDate + vbCrLf
    tmp = tmp + "errorCode(오류코드) : " + CStr(jobInfo.errorCode) + vbCrLf
    tmp = tmp + "errorReason(오류메시지) : " + jobInfo.errorReason + vbCrLf
    tmp = tmp + "jobStartDT(작업 시작일시) : " + jobInfo.jobStartDT + vbCrLf
    tmp = tmp + "jobEndDT(작업 종료일시) : " + jobInfo.jobEndDT + vbCrLf
    tmp = tmp + "regDT(수집 요청일시) : " + jobInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = easyFinBankService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = easyFinBankService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
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
    joinData.id = "userid"
    
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
    
    Set Response = easyFinBankService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' RequestJob(수집 요청)를 통해 반환 받은 작업아이디의 목록을 확인합니다.
' - 수집 요청 후 1시간이 경과한 수집 요청건은 상태정보가 반환되지 않습니다.
' - https://docs.popbill.com/easyfinbank/vb/api#ListActiveJob
'=========================================================================
Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
    Dim tmp As String
    Dim info As PBEasyFinBankJobState
    
    Set jobList = easyFinBankService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "작업아이디(jobID)의 유효시간은 1시간입니다" + vbCrLf + vbCrLf
    tmp = tmp + "jobID(작업아이디) | jobState(수집상태) | startDate(시작일자) | endDate(종료일자) |" _
            + "errorCode(오류코드) | errorReason(오류메시지) | jobStartDT(작업 시작일시) | jobEndDT(작업 종료일시) | regDT(수집 요청일시) " + vbCrLf
    
    For Each info In jobList
        tmp = tmp + CStr(info.jobID) + " | "
        tmp = tmp + CStr(info.jobState) + " | "
        tmp = tmp + info.startDate + " | "
        tmp = tmp + info.endDate + " | "
        tmp = tmp + CStr(info.errorCode) + " | "
        tmp = tmp + info.errorReason + " | "
        tmp = tmp + info.jobStartDT + " | "
        tmp = tmp + info.jobEndDT + " | "
        tmp = tmp + info.regDT
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
    If jobList.count > 0 Then
        txtJobID.Text = jobList.Item(1).jobID
    End If
End Sub

'=========================================================================
' 팝빌에 등록된 은행계좌 목록을 반환한다.
' - https://docs.popbill.com/easyfinbank/vb/api#ListBankAccount
'=========================================================================
Private Sub btnListBankAccount_Click()
    Dim bankAccountList As Collection
    Dim tmp As String
    Dim info As PBEasyFinBankAccount
    
    Set bankAccountList = easyFinBankService.ListBankAccount(txtCorpNum.Text)
     
    If bankAccountList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    
    tmp = "accountNumber(계좌번호) | bankCode(은행코드) | accountName(계좌별칭) | accountType(계좌유형) | state(계좌 상태) | regDT(등록일시) |" _
        + " contractState (정액제 서비스 상태) | closeRequestYN (정액제 해지신청 여부) | useRestrictYN (정액제 사용제한 여부) | closeOnExpired (정액제 만료시 해지여부) | " _
        + " unPaiedYN (미수금 보유 여부) | memo(메모)" + vbCrLf + vbCrLf
    
    For Each info In bankAccountList
        tmp = tmp + info.AccountNumber + " | "
        tmp = tmp + info.BankCode + " | "
        tmp = tmp + info.AccountName + " | "
        tmp = tmp + info.AccountType + " | "
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.regDT + " | "
        
        tmp = tmp + info.contractDT + " | "
        tmp = tmp + CStr(info.baseDate) + " | "
        tmp = tmp + info.useEndDate + " | "
        tmp = tmp + CStr(info.contractState) + " | "
        tmp = tmp + CStr(info.closeRequestYN) + " | "
        tmp = tmp + CStr(info.useRestrictYN) + " | "
        tmp = tmp + CStr(info.closeOnExpired) + " | "
        tmp = tmp + CStr(info.unPaidYN) + " | "
        tmp = tmp + info.Memo
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = easyFinBankService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchAllAllowYN(회사조회 권한여부) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 계좌조회 서비스를 이용할 계좌를 팝빌에 등록합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#RegistBankAccount
'=========================================================================
Private Sub btnRegistBankAccount_Click()
    Dim AccountInfo As New PBEasyFinBankAccountForm
    Dim Response As PBResponse
    
    ' [필수] 은행코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    AccountInfo.BankCode = ""
    
    ' [필수] 계좌번호 하이픈('-') 제외
    AccountInfo.AccountNumber = ""

    ' [필수] 계좌비밀번호
    AccountInfo.AccountPWD = ""

    ' [필수] 계좌유형, "법인" 또는 "개인" 입력
    AccountInfo.AccountType = ""

    ' [필수] 예금주 식별정보 (‘-‘ 제외)
    ' 계좌유형이 “법인”인 경우 : 사업자번호(10자리)
    ' 계좌유형이 “개인”인 경우 : 예금주 생년월일 (6자리-YYMMDD)
    AccountInfo.IdentityNumber = ""

    ' 계좌 별칭
    AccountInfo.AccountName = ""

    ' 인터넷뱅킹 아이디 (국민은행 필수)
    AccountInfo.BankID = ""

    ' 조회전용 계정 아이디 (대구은행, 신협, 신한은행 필수)
    AccountInfo.FastID = ""

    ' 조회전용 계정 비밀번호 (대구은행, 신협, 신한은행 필수
    AccountInfo.FastPWD = ""

    ' 결제기간(개월), 1~12 입력가능, 미기재시 기본값(1) 처리
    ' - 파트너 과금방식의 경우 입력값에 관계없이 1개월 처리
    AccountInfo.UsePeriod = "1"

    ' 메모
    AccountInfo.Memo = ""
    
   
    Set Response = easyFinBankService.RegistBankAccount(txtCorpNum.Text, AccountInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.pwd = "test@test.com"
    
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
    
    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
        
    Set Response = easyFinBankService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 계좌 거래내역을 확인하기 위해 팝빌에 수집요청을 합니다. 조회기간은 당일 기준으로 90일 이내로만 지정 가능합니다.
' - 반환 받은 작업아이디는 함수 호출 시점부터 1시간 동안 유효합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#RequestJob
'=========================================================================

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim BankCode As String
    Dim AccountNumber As String
    Dim SDate As String
    Dim EDate As String
    
    '은행코드
    BankCode = "0004"
    
    '팝빌에 등록된 계좌번호
    AccountNumber = "20700644024204"
    
    '시작일자, 표시형식(yyyyMMdd)
    SDate = "20210901"
    
    '종료일자, 표시형식(yyyyMMdd)
    EDate = "20210910"
    
    jobID = easyFinBankService.RequestJob(txtCorpNum.Text, BankCode, AccountNumber, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(작업아이디) : " + jobID + vbCrLf
    
    txtJobID.Text = jobID
End Sub

'=========================================================================
' 신청한 정액제 해지요청을 취소합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#RevokeCloseBankAccount
'=========================================================================
Private Sub btnRevokeCloseBankAccount_Click()

    Dim Response As PBResponse
    Dim BankCode As String
    Dim AccountNumber As String
    
    ' [필수] 은행코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    BankCode = ""
    
    ' [필수] 계좌번호 하이픈('-') 제외
    AccountNumber = ""
    
    Set Response = easyFinBankService.RevokeCloseBankAccount(txtCorpNum.Text, BankCode, AccountNumber)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 한 건의 거래 내역에 메모를 저장합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#SaveMemo
'=========================================================================
Private Sub btnSaveMemo_Click()
    Dim Response As PBResponse
    Dim tid As String
    Dim Memo As String
    
    ' 거래내역 아이디, SeachAPI 응답항목 중 tid
    tid = "02112181100000000120211210000003"
    
    '메모
    Memo = "메모 테스트"
    
    Set Response = easyFinBankService.SaveMemo(txtCorpNum.Text, tid, Memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub


'=========================================================================
' GetJobState(수집 상태 확인)를 통해 상태 정보가 확인된 작업아이디를 활용하여 계좌 거래 내역을 조회합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#Search
'=========================================================================
Private Sub btnSearch_Click()
    Dim searchList As PBEasyFinBankSearch
    Dim TradeType As New Collection
    Dim page As Integer
    Dim perPage As Integer
    Dim order As String
    Dim tmp As String
    Dim listboxRow As String
    Dim SearchString As String
        
    '거래유형 배열, I-입금, O-지출
    TradeType.Add "I"
    TradeType.Add "O"
    
    '페이지번호, 기본값 ‘1’
    page = 1
    
    '페이지당 검색개수, 기본값 500, 최대 1000
    perPage = 20
    
    '정렬 방향, D-내림차순, A-오름차순
    order = "D"
    
    '조회 검색어, 입금/출금액, 메모, 적요 like 검색
    SearchString = ""
        
    Set searchList = easyFinBankService.Search(txtCorpNum.Text, txtJobID.Text, TradeType, SearchString, page, perPage, order, txtUserID.Text)
    
        
    If searchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(searchList.code) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + searchList.Message + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(searchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(searchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(searchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(searchList.pageCount) + vbCrLf
    tmp = tmp + "lastScrapDT (최종 조회일시) : " + searchList.lastScrapDT + vbCrLf + vbCrLf
    
    searchInfo.Clear
        
    searchInfo.AddItem "tid (거래내역 아이디) | trdate (거래일자) | trserial (거래일자별 거래내역순번) | trdt (거래일시) | accIn (입금액) | accOut (출금액) | balance (잔액) | ", 0
    searchInfo.AddItem "remark1 (비고) | remark2 (비고) | remark3 (비고) | remark4 (비고) | regDT (등록일시) | memo (메모)", 1
    
    Dim tiInfo As PBEasyFinBankSearchDetail
           
    For Each tiInfo In searchList.list
        listboxRow = ""
        listboxRow = tiInfo.tid + " | "
        listboxRow = listboxRow + tiInfo.trdate + " | "
        listboxRow = listboxRow + CStr(tiInfo.trserial) + " | "
        listboxRow = listboxRow + tiInfo.trdt + " | "
        listboxRow = listboxRow + tiInfo.accIn + " | "
        listboxRow = listboxRow + tiInfo.accOut + " | "
        listboxRow = listboxRow + tiInfo.balance + " | "
        listboxRow = listboxRow + tiInfo.remark1 + " | "
        listboxRow = listboxRow + tiInfo.remark2 + " | "
        listboxRow = listboxRow + tiInfo.remark3 + " | "
        listboxRow = listboxRow + tiInfo.remark4 + " | "
        listboxRow = listboxRow + tiInfo.regDT + " | "
        listboxRow = listboxRow + tiInfo.Memo
        searchInfo.AddItem listboxRow, searchInfo.ListCount
    Next
  
    MsgBox (tmp)
End Sub

'=========================================================================
' GetJobState(수집 상태 확인)를 통해 상태 정보가 확인된 작업아이디를 활용하여 계좌 거래내역의 요약 정보를 조회합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#Summary
'=========================================================================
Private Sub btnSummary_Click()
    Dim summaryInfo As PBEasyFinBankSummary
    Dim TradeType As New Collection
    Dim SearchString As String
    Dim tmp As String
    
    '거래유형 배열, I-입금, O-지출
    TradeType.Add "I"
    TradeType.Add "O"
    
    '조회 검색어, 입금/출금액, 메모, 적요 like 검색
    SearchString = ""
    
    Set summaryInfo = easyFinBankService.Summary(txtCorpNum.Text, txtJobID.Text, TradeType, SearchString, txtUserID.Text)
        
    If summaryInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (수집결과 건수) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "cntAccIn (입금거래 건수) : " + CStr(summaryInfo.cntAccIn) + vbCrLf
    tmp = tmp + "cntAccOut (출금거래 건수) : " + CStr(summaryInfo.cntAccOut) + vbCrLf
    tmp = tmp + "totalAccIn (입금액 합계) : " + CStr(summaryInfo.totalAccIn) + vbCrLf
    tmp = tmp + "totalAccOut (출금액 합계) : " + CStr(summaryInfo.totalAccOut) + vbCrLf
       
    MsgBox (tmp)
End Sub

'=========================================================================
' 팝빌에 등록된 계좌정보를 수정합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateBankAccount
'=========================================================================
Private Sub btnUpdateBankAccount_Click()
    Dim AccountInfo As New PBEasyFinBankAccountForm
    Dim Response As PBResponse
    
    ' [필수] 은행코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    AccountInfo.BankCode = ""
    
    ' [필수] 계좌번호 하이픈('-') 제외
    AccountInfo.AccountNumber = ""

    ' [필수] 계좌비밀번호
    AccountInfo.AccountPWD = ""

    ' 계좌 별칭
    AccountInfo.AccountName = ""

    ' 인터넷뱅킹 아이디 (국민은행 필수)
    AccountInfo.BankID = ""

    ' 조회전용 계정 아이디 (대구은행, 신협, 신한은행 필수)
    AccountInfo.FastID = ""

    ' 조회전용 계정 비밀번호 (대구은행, 신협, 신한은행 필수
    AccountInfo.FastPWD = ""
    
    ' 메모
    AccountInfo.Memo = ""
    
   
    Set Response = easyFinBankService.UpdateBankAccount(txtCorpNum.Text, AccountInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
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

    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
                
    Set Response = easyFinBankService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateCorpInfo
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
    
    Set Response = easyFinBankService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

Private Sub Form_Load()

    '간편 계좌조회 서비스 초기화
    easyFinBankService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(개발용), False(상업용)
    easyFinBankService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True(권장)
    easyFinBankService.IPRestrictOnOff = True
    
End Sub

