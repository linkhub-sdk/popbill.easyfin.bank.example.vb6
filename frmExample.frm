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
         Left            =   7680
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
         Width           =   2295
         Begin VB.CommandButton btnListBankAccount 
            Caption         =   "계좌 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBankAccountMgtURL 
            Caption         =   "계좌관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   360
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
' 팝빌 간편 계좌조회 API VB 6.0 SDK Example
'
' - 업데이트 일자 : 2019-12-20
' - 연동 기술지원 연락처 : 1600-8536 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
'=========================================================================

Option Explicit

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'홈택스 전자세금계산서 연동 클래스 선언
Private easyFinBankService As New PBEasyFinBankService

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
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
' 파트너의 연동회원으로 가입된 사업자번호인지 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
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
' 팝빌에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
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
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
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
' 은행 계좌관리 팝업 URL을 확인합니다.
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
' 연동회원의 간편 계좌조회 API 서비스 과금정보를 확인합니다.
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
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
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
' 간편 계좌조회 정액제 서비스 신청 팝업 URL을 반환한다.
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
' 정액제 서비스 상태를 확인합니다.
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBEasyFinBankFlatRate
    Dim tmp As String
    Dim bankCode As String
    Dim accountNumber As String
    
    '은행코드
    bankCode = "0048"
    
    '팝빌에 등록된 계좌번호
    accountNumber = "131020538645"
    
    Set flatRateInfo = easyFinBankService.GetFlatRateState(txtCorpNum.Text, bankCode, accountNumber)
     
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
' 계좌 거래내역 수집 상태를 확인한다.
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
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)
'   를 통해 확인하시기 바랍니다.
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
' 파트너 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
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
' 팝빌 연동회원 가입을 요청합니다.
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
' 1시간 이내 수집 요청 목록을 확인한다.
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
            + "memo(메모)" + vbCrLf
    
    For Each info In bankAccountList
        tmp = tmp + info.accountNumber + " | "
        tmp = tmp + info.bankCode + " | "
        tmp = tmp + info.accountName + " | "
        tmp = tmp + info.accountType + " | "
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.regDT + " | "
        tmp = tmp + info.memo
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
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
' 연동회원의 담당자를 신규로 등록합니다.
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
' 계좌 거래내역 수집을 요청한다.
' - 검색기간은 현재일 기준 90일 이내로만 요청할 수 있다.
'=========================================================================

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim bankCode As String
    Dim accountNumber As String
    Dim SDate As String
    Dim EDate As String
    
    '은행코드
    bankCode = "0004"
    
    '팝빌에 등록된 계좌번호
    accountNumber = "131020538600"
    
    '시작일자, 표시형식(yyyyMMdd)
    SDate = "20190921"
    
    '종료일자, 표시형식(yyyyMMdd)
    EDate = "20191220"
    
    jobID = easyFinBankService.RequestJob(txtCorpNum.Text, bankCode, accountNumber, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(작업아이디) : " + jobID + vbCrLf
    
    txtJobID.Text = jobID
End Sub

'=========================================================================
' 계좌 거래내역에 메모를 저장한다.
'=========================================================================
Private Sub btnSaveMemo_Click()
    Dim Response As PBResponse
    Dim tid As String
    Dim memo As String
    
    ' 거래내역 아이디, SeachAPI 응답항목 중 tid
    tid = "01912181100000000120191210000003"
    
    '메모
    memo = "20191220-테스트"
    
    Set Response = easyFinBankService.SaveMemo(txtCorpNum.Text, tid, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "응답메시지 : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub


'=========================================================================
' 계좌 거래내역을 조회한다.
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
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(searchList.pageCount) + vbCrLf + vbCrLf
    
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
        listboxRow = listboxRow + tiInfo.memo
        searchInfo.AddItem listboxRow, searchInfo.ListCount
    Next
  
    MsgBox (tmp)
End Sub

'=========================================================================
' 계좌 거래내역 요약정보를 조회한다.
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
' 연동회원의 담당자 정보를 수정합니다.
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
' 연동회원의 회사정보를 수정합니다
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

