VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSCreditCardCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'==============================================================================
' YOU ARE FREE TO USE THIS CODE IN YOUR OWN VB PROJECTS PROVIDED NO
' CHANGES ARE MADE TO THE ORIGINAL SOURCE CODE. PLEASE REPORT BUGS
' TO const71@yahoo.com
' Copyright(c) 2003 Constantin Nterekas
' http://www.foundationssoftware.com
'==============================================================================

Public Enum CreditCardConstants
    AMERICAN_EXPRESS = 1
    DINERS_CLUB = 2
    DISCOVER = 3
    JCB = 4
    MASTERCARD = 5
    VISA = 6
End Enum

'==============================================================================
'   PROCEDURE: VerifyNumber
' DESCRIPTION: Validates the specified account number and expiry date.
'  PARAMETERS:  AccountNumber - the Account Number to be validated.
'              CreditCardType - the CreditCardConstant indicating the
'                               credit card's type.
'                 ExpiryMonth - the 1 or 2 digit number representing the month.
'                  ExpiryYear - the 4 digit number representing the year.
'               ResultMessage - (Call By Reference) clients may use this for a
'                               description of the success or failure
'                               of verification.
'     RETURNS: Boolean - True if the Account and Expiry Date are Valid. If
'                        False is returned, clients may look at the
'                        ResultMessage parameter to uncover the reasons why
'                        verification failed.
'==============================================================================
Public Function VerifyNumber(ByVal AccountNumber As String, _
                       Optional ByVal CreditCardType As CreditCardConstants = 0, _
                       Optional ByVal ExpiryMonth As Long = 0, _
                       Optional ByVal ExpiryYear As Long = 0, _
                       Optional ByRef ResultMessage As String = "") As Boolean
                       
    Dim mDiscoveredType As CreditCardConstants
      
    VerifyNumber = True
    ' MAKE SURE ALL CHARACTERS ARE DIGITS
    If Not (IsNumber(AccountNumber)) Then
        ResultMessage = "The account " & AccountNumber & " contains illegal characters." & vbCrLf
        ResultMessage = ResultMessage & "It must be free of all spaces and non-numeric characters."
        VerifyNumber = False
        Exit Function
    End If
      
    ' MAKE SURE PREFIX EXISTS AND TAKE NOTE OF THE DISCOVERED TYPE.
    ' LENGTHS MUST ALSO BE IN ORDER ...
    mDiscoveredType = CheckPrefixAndLength(AccountNumber)
    If (mDiscoveredType = 0) Then
        ResultMessage = "The account " & AccountNumber & " cannot be matched with "
        ResultMessage = ResultMessage & "a known credit card type."
        VerifyNumber = False
        Exit Function
    End If
    
    ' IF THE USER SUPPLIED US WITH THE OPTIONAL CreditCardType PARAMETER, WE MUST
    ' MAKE SURE IT IS THE SAME TYPE THAT WE JUST DISCOVERED IT TO BE ...
    If (CreditCardType <> 0) Then
        If (mDiscoveredType <> CreditCardType) Then
            ResultMessage = "The supplied credit card type (" & GetCardTypeDescr(CreditCardType) & ") is incorrect." & vbCrLf
            ResultMessage = ResultMessage & "The account " & AccountNumber & " has a format that may be" & vbCrLf
            ResultMessage = ResultMessage & "consistent with " & GetCardTypeDescr(mDiscoveredType) & " types."
            VerifyNumber = False
            Exit Function
        End If
    End If
    
    'PERFORM LUHN (MOD 10) VALIDATION ...
    If Not (LUHNMod10Verify(AccountNumber)) Then
        ResultMessage = "The account number " & AccountNumber & " does not appear to be a valid" & vbCrLf
        ResultMessage = ResultMessage & GetCardTypeDescr(mDiscoveredType) & ". Double check the "
        ResultMessage = ResultMessage & "number and try again."
        VerifyNumber = False
        Exit Function
    End If
    
    'VALIDATE THE EXPIRY DATE IF IT WAS PROVIDED
    If (ExpiryMonth <> 0 Or ExpiryYear <> 0) Then
        If Not (DateVerify(ExpiryMonth, ExpiryYear)) Then
            ResultMessage = "The expiry date is incorrect."
            VerifyNumber = False
            Exit Function
        End If
    End If
    
    'RETURN A SUCCESSFUL MESSAGE TO USER ...
    ResultMessage = "Valid " & GetCardTypeDescr(mDiscoveredType)
End Function
'==============================================================================
'   PROCEDURE: CleanAccountString
' DESCRIPTION: Strips all non-numeric characters including spaces from the
'              parameter OriginalAccount and returns the result as a String.
'  PARAMETERS: OriginalAccount - the Account Number to clean.
'     RETURNS: String - the cleaned Account Number.
'==============================================================================
Public Function CleanAccountString(ByVal OriginalAccount As String) As String
    Dim mResult As String
    Dim mIndex As Long

    mResult = ""
    For mIndex = 1 To Len(OriginalAccount)
        If (IsNumber(Mid(OriginalAccount, mIndex, 1))) Then
            mResult = mResult & Mid(OriginalAccount, mIndex, 1)
        End If
    Next mIndex
    CleanAccountString = mResult
End Function
'==============================================================================
'   PROCEDURE: GetCreditCardType
' DESCRIPTION: Returns the CreditCardConstant representing the credit card type.
'  PARAMETERS: AccountNumber - the AccountNumber to get the credit card type for.
'     RETURNS: The CreditCardConstant representing the credit card type. Zero
'              is returned if a valid type cannot be determined.
'==============================================================================
Public Function GetCreditCardType(ByVal AccountNumber As String) As Long
    GetCreditCardType = 0
    If (VerifyNumber(AccountNumber, CreditCardConstants.AMERICAN_EXPRESS)) Then
        GetCreditCardType = CreditCardConstants.AMERICAN_EXPRESS
    ElseIf (VerifyNumber(AccountNumber, CreditCardConstants.DINERS_CLUB)) Then
        GetCreditCardType = CreditCardConstants.DINERS_CLUB
    ElseIf (VerifyNumber(AccountNumber, CreditCardConstants.DISCOVER)) Then
        GetCreditCardType = CreditCardConstants.DISCOVER
    ElseIf (VerifyNumber(AccountNumber, CreditCardConstants.JCB)) Then
        GetCreditCardType = CreditCardConstants.JCB
    ElseIf (VerifyNumber(AccountNumber, CreditCardConstants.MASTERCARD)) Then
        GetCreditCardType = CreditCardConstants.MASTERCARD
    Else
        If (VerifyNumber(AccountNumber, CreditCardConstants.VISA)) Then
        GetCreditCardType = CreditCardConstants.VISA
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: CheckPrefixAndLength
' DESCRIPTION: This helper function checks the prefix and length of the account.
'              It assumes the account being checked is comprised of numbers only.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Long - the CreditCardConstant that maps to this account or
'              zero if no match can be made.
'==============================================================================
Private Function CheckPrefixAndLength(ByVal AccountNumber As String) As Long
    Dim mPrefix As Long

    CheckPrefixAndLength = 0
    ' VISA HAS SHORTEST ALLOWABLE LENGTH OF 13 AND SEVERAL TYPES USE 16
    ' SO LET'S ELIMINATE BOGUS LENGTHS RIGHT OF THE BAT...
    If (Len(AccountNumber) >= 13 And Len(AccountNumber) <= 16) Then
        If (IsAmericanExpress(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.AMERICAN_EXPRESS:    Exit Function
        If (IsDinersClub(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.DINERS_CLUB: Exit Function
        If (IsDiscover(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.DISCOVER: Exit Function
        If (IsJCB(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.JCB: Exit Function
        If (IsMasterCard(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.MASTERCARD:    Exit Function
        If (IsVisa(AccountNumber)) Then CheckPrefixAndLength = CreditCardConstants.VISA:    Exit Function
    End If
End Function
'==============================================================================
'   PROCEDURE: IsAmericanExpress
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              American Express card. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        American Express Card. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsAmericanExpress(ByVal AccountNumber As String) As Boolean
    IsAmericanExpress = False
    If (Len(Trim(AccountNumber)) = 15) Then
        If (Left(AccountNumber, 2) = "34" Or Left(AccountNumber, 2) = "37") Then
            IsAmericanExpress = True
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: IsDinersClub
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              Diners Club card. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        Diners Club Card. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsDinersClub(ByVal AccountNumber As String) As Boolean
    IsDinersClub = False
    If (Len(Trim(AccountNumber)) = 14) Then
        If (Left(AccountNumber, 3) = "300" Or _
            Left(AccountNumber, 3) = "301" Or _
            Left(AccountNumber, 3) = "302" Or _
            Left(AccountNumber, 3) = "303" Or _
            Left(AccountNumber, 3) = "304" Or _
            Left(AccountNumber, 3) = "305" Or _
            Left(AccountNumber, 2) = "36" Or _
            Left(AccountNumber, 2) = "38") Then
            IsDinersClub = True
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: IsDiscover
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              Discover card. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        Discover Card. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsDiscover(ByVal AccountNumber As String) As Boolean
    IsDiscover = False
    If (Len(Trim(AccountNumber)) = 16) Then
        If (Left(AccountNumber, 4) = "6011") Then
            IsDiscover = True
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: IsJCB
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              JCB card. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        JCB Card. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsJCB(ByVal AccountNumber As String) As Boolean
    IsJCB = False
    If (Len(Trim(AccountNumber)) = 16) Then
        If (Left(AccountNumber, 1) = "3") Then
            IsJCB = True
        End If
    Else
        If (Len(Trim(AccountNumber)) = 15) Then
            If (Left(AccountNumber, 4) = "2131" Or Left(AccountNumber, 4) = "1800") Then
                IsJCB = True
            End If
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: IsMasterCard
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              MasterCard. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        MasterCard. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsMasterCard(ByVal AccountNumber As String) As Boolean
    IsMasterCard = False
    If (Len(Trim(AccountNumber)) = 16) Then
        If (Left(AccountNumber, 2) = "51" Or _
            Left(AccountNumber, 2) = "52" Or _
            Left(AccountNumber, 2) = "53" Or _
            Left(AccountNumber, 2) = "54" Or _
            Left(AccountNumber, 2) = "55") Then
            IsMasterCard = True
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: IsVisa
' DESCRIPTION: This helper function determines whether the prefix and length of
'              the specified account qualify this account as a possible
'              Visa card. It assumes the length of the passed
'              parameter is between 13 and 16 chars and consists exclusively
'              of numbers.
'  PARAMETERS: AccountNumber - the account number to check
'     RETURNS: Boolean - true if this candidate account can pass for a valid
'                        Visa Card. We can't be certain of this
'                        until after a MOD 10 validation test but this is not
'                        performed here ...
'==============================================================================
Private Function IsVisa(ByVal AccountNumber As String) As Boolean
    IsVisa = False
    If (Len(Trim(AccountNumber)) = 13 Or Len(Trim(AccountNumber)) = 16) Then
        If (Left(AccountNumber, 1) = "4") Then
            IsVisa = True
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: GetCardTypeDescr
' DESCRIPTION: This function returns the credit card types description given
'              the specified CreditCardConstant.
'  PARAMETERS: CardType - the constant whose descriptive name is returned.
'     RETURNS: String - The credit card type description or an empty String if
'                       no type was found.
'==============================================================================
Public Function GetCardTypeDescr(ByVal CardType As CreditCardConstants) As String
    GetCardTypeDescr = ""
    If (CardType = CreditCardConstants.AMERICAN_EXPRESS) Then
        GetCardTypeDescr = "American Express"
    ElseIf (CardType = CreditCardConstants.DINERS_CLUB) Then
        GetCardTypeDescr = "Diners Club"
    ElseIf (CardType = CreditCardConstants.DISCOVER) Then
        GetCardTypeDescr = "Discover"
    ElseIf (CardType = CreditCardConstants.JCB) Then
        GetCardTypeDescr = "JCB"
    ElseIf (CardType = CreditCardConstants.MASTERCARD) Then
        GetCardTypeDescr = "Master Card"
    Else
        If (CardType = CreditCardConstants.VISA) Then
            GetCardTypeDescr = "Visa"
        End If
    End If
End Function
'==============================================================================
'   PROCEDURE: LUHNMod10Verify
' DESCRIPTION: This helper function applies the LUHN formula (MOD 10) to determine
'              whether this number is valid or not. This function assumes
'              the length, prefix and contents of AccountNumber are otherwise
'              correct. For more information about this formula, refer to
'              README.TXT
'  PARAMETERS: AccountNumber - the number to check.
'     RETURNS: Boolean - true if all is ok or false if the validtion failed.
'==============================================================================
Private Function LUHNMod10Verify(ByVal AccountNumber As String) As Boolean
    Dim mResult As String
    Dim mChar As String
    Dim mPosition As Long
    Dim mSum As Long

    ' FIRST WE REVERSE THE STRING SO WE CAN WORK FROM LEFT TO RIGHT
    ' WHICH IS MORE NATURAL TO FOLLOW ...
    mResult = Trim(StrReverse(AccountNumber))
    ' NEXT, WE LOOP THROUGH EACH NUMERIC CHARACTER AND PERFORM DIFFERENT
    ' CALCULATIONS FOR ODD AND EVEN POSITIONS ...
    mSum = 0
    For mPosition = 1 To Len(mResult)
        ' GRAB THE CURRENT DIGIT ...
        mChar = Mid(mResult, mPosition, 1)
                
        If (mPosition Mod 2 = 0) Then
           ' THIS IS AN EVEN POSITION SO FIRST DOUBLE UP THIS DIGIT AND THEN
           ' ADD THE INDIVIDUAL DIGITS FROM THE RESULT. THE INDIVIDUAL DIGITS
           ' ARE SUMMED BY THE FUNCTION AddDigits..
           mSum = mSum + AddDigits(CStr(CLng(mChar) * 2))
        Else
           ' THIS IS AN ODD POSITION SO SIMPLY ADD IT AS IS ...
           mSum = mSum + CLng(mChar)
        End If
    Next mPosition
    ' mSum MUST BE EVENLY DIVISIBLE BY 10, OTHERWISE THE ACCOUNT IS NOT VALID...
    LUHNMod10Verify = (mSum Mod 10 = 0)
End Function
'==============================================================================
'   PROCEDURE: AddDigits
' DESCRIPTION: This helper function adds up the individual digits that make up
'              the specified value (mValue). This function assumes the passed
'              value is comprised entirely of numbers. For example:
'
'    mValue   Result  Comments
'    ======   ======  ========
'       "1"        1  no change
'       "2"        2  no change
'       "3"        3  no change
'       "9"        9  no change
'      "10"        1  (1 + 0)
'      "11"        2  (1 + 1)
'      "12"        3  (2 + 1)
'      "18"        9  (8 + 1)
'
'  PARAMETERS: mValue - the number to sum.
'     RETURNS: Long - the result after the addition is performed.
'==============================================================================
Private Function AddDigits(ByVal mValue As String) As Long
    Dim iPosition As Long
    Dim mSum As Long
    
    mSum = 0
    For iPosition = 1 To Len(mValue)
        mSum = mSum + CLng(Mid(mValue, iPosition, 1))
    Next iPosition
    AddDigits = mSum
End Function
'==============================================================================
'   PROCEDURE: DateVerify
' DESCRIPTION: This helper function validates the expiry year and month.
'  PARAMETERS: ExpiryMonth - the month to check.
'               ExpiryYear - the year to check.
'     RETURNS: Boolean - true if all is ok or false if the validation failed.
'==============================================================================
Private Function DateVerify(ByVal ExpiryMonth As Long, ByVal ExpiryYear As Long) As Boolean
    Dim mYear As Long
    Dim mMonth As Long

    DateVerify = True
    If (ExpiryYear <> 0) Then
        ' GET TODAYS YEAR AND MAKE SURE THE EXPIRY YEAR HASN'T EXPIRED
        ' ALSO, MAKE SURE THE EXPIRY YEAR ISNT MORE THAN 100 YRS FROM NOW
        mYear = CLng(DatePart("yyyy", Now))
        If (ExpiryYear < mYear) Then DateVerify = False: Exit Function
        If (ExpiryYear > (mYear + 100)) Then DateVerify = False: Exit Function
        ' WE NOW EXAMINE THE MONTH IF IT IS GREATER THAN ZERO AND VALIDATE IT
        ' AS WELL...
        If (ExpiryMonth <> 0) Then
            ' FIRST MAKE SURE THE MONTH IS BETWEEN 1 AND 12 ...
            If (ExpiryMonth < 1) Then DateVerify = False: Exit Function
            If (ExpiryMonth > 12) Then DateVerify = False: Exit Function
            ' SINCE WE KNOW THE EXPIRY YEAR IS OK AT THIS POINT, LETS MAKE
            ' SURE THE EXPIRY MONTH HAS NOT PASSED IF THE CURRENT YEAR AND
            ' EXPIRY YEAR ARE THE SAME ...
            If (mYear = ExpiryYear) Then
                ' EXPIRES THIS YEAR...DOUBLE CHECK THE MONTH HASN'T EXPIRED ...
                mMonth = CLng(DatePart("m", Now))
                If (ExpiryMonth < mMonth) Then DateVerify = False: Exit Function
            End If
        End If
    End If
    
    ' WE NEED TO CHECK THE EXPIRY MONTH IN ISOLATION (I.E. YEAR WAS OMITTED)...
    If (ExpiryMonth <> 0) Then
        ' NO NEED TO TAKE THE YEAR INTO ACCOUNT HERE SINCE WE VALIDATED
        ' THAT IN LAST STEP. JUST MAKE SURE THE RANGE IS OK.
        If (ExpiryMonth < 1) Then DateVerify = False: Exit Function
        If (ExpiryMonth > 12) Then DateVerify = False: Exit Function
    End If
End Function
