VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card Demo"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox YearField 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3630
      MaxLength       =   4
      TabIndex        =   7
      Top             =   870
      Width           =   1755
   End
   Begin VB.TextBox MonthField 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2610
      MaxLength       =   2
      TabIndex        =   5
      Top             =   870
      Width           =   975
   End
   Begin VB.ComboBox CardTypes 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   2775
   End
   Begin VB.CommandButton CheckButton 
      Caption         =   "Validate"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   1290
      Width           =   5265
   End
   Begin VB.TextBox AccountNumber 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2610
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Expiry Date (MM-YYYY):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   2445
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Card Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   495
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' A SIMPLE DEMO USING THE CREDIT CARD VALIDATOR DLL...
'
' YOU ARE FREE TO USE THIS CODE IN YOUR OWN VB PROJECTS PROVIDED NO
' CHANGES ARE MADE TO THE ORIGINAL SOURCE CODE. PLEASE REPORT BUGS
' TO const71@yahoo.com
' Copyright(c) 2003 Constantin Nterekas
' http://www.foundationssoftware.com
'==============================================================================

Private Sub Form_Load()
    CardTypes.AddItem (""): CardTypes.ItemData(CardTypes.NewIndex) = 0
    CardTypes.AddItem ("American Express"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.AMERICAN_EXPRESS
    CardTypes.AddItem ("Diners Club"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.DINERS_CLUB
    CardTypes.AddItem ("Discover"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.DISCOVER
    CardTypes.AddItem ("JCB"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.JCB
    CardTypes.AddItem ("MasterCard"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.MASTERCARD
    CardTypes.AddItem ("Visa"): CardTypes.ItemData(CardTypes.NewIndex) = CreditCardConstants.VISA
    CardTypes.ListIndex = 0
End Sub

Private Sub CheckButton_Click()
    Dim checker As FSCreditCardCheck
    Dim msg As String
    Dim mMonth As Long
    Dim mYear As Long

    mMonth = 0
    mYear = 0
    
    If (IsNumeric(MonthField)) Then
        mMonth = CLng(MonthField.Text)
    End If

    If (IsNumeric(YearField)) Then
        mYear = CLng(YearField.Text)
    End If

    Set checker = New FSCreditCardCheck
    AccountNumber.Text = Trim(checker.CleanAccountString(AccountNumber.Text))
    If (checker.VerifyNumber(AccountNumber.Text, _
                             CardTypes.ItemData(CardTypes.ListIndex), _
                             mMonth, _
                             mYear, _
                             msg)) Then
        MsgBox "Card is valid"
    Else
        MsgBox "Card is invalid"
    End If
    MsgBox msg
    MsgBox checker.GetCardTypeDescr(checker.GetCreditCardType(AccountNumber.Text))
    Set checker = Nothing
End Sub

