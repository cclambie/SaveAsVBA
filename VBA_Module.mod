Option Explicit
Sub ExportAsCSV()

    Dim MyFileName As String
    Dim CurrentWB As Workbook, TempWB As Workbook

    Set CurrentWB = ActiveWorkbook
    ActiveWorkbook.ActiveSheet.UsedRange.Copy

    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    
    'Dim Change below to "- 4"  to become compatible with .xls files
    MyFileName = CurrentWB.Path & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & ".csv"

    Application.DisplayAlerts = False
    TempWB.SaveAs Filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub


Public Sub cbGenerateOFX()
' Generate OFX file from the data in the 'Export to OFX' spreadsheet

On Error Resume Next

' General variables
Dim iResult As Integer
Dim iTransaction As Integer
Dim dtDatetime As Date

' Output file name
Dim OutputFilename As String

Dim CurrentWB As Workbook, TempWB As Workbook

    Set CurrentWB = ActiveWorkbook

' OFX file header
Dim OFX_HEADER As String
Dim OFX_SIGNONMSGSRSV1_HEADER As String
Dim OFX_SIGNONMSGSRSV1_DTSERVER As String
Dim OFX_SIGNONMSGSRSV1_FOOTER As String
Dim OFX_BANKMSGSRSV1_HEADER As String
Dim OFX_BANKMSGSRSV1_FOOTER As String

' Bank account information
Dim OFX_BANKACCTFROM_HEADER As String
Dim OFX_BANKID As String
Dim OFX_ACCTID As String
Dim OFX_ACCTTYPE As String
Dim OFX_BANKACCTFROM_FOOTER As String

' Transaction list information
Dim OFX_BANKTRANLIST_HEADER As String
Dim OFX_BANKTRANLIST_DTSTART As String
Dim OFX_BANKTRANLIST_DTEND As String
Dim OFX_BANKTRANLIST_FOOTER As String

' Transactions information
Dim OFX_STMTTRN_HEADER As String
Dim OFX_STMTTRN_TRNTYPE As String
Dim OFX_STMTTRN_DTPOSTED As String
Dim OFX_STMTTRN_TRNAMT As String
Dim OFX_STMTTRN_FITID As String
Dim OFX_STMTTRN_NAME As String
Dim OFX_STMTTRN_MEMO As String
Dim OFX_STMTTRN_FOOTER As String

' Ledger balance information
Dim OFX_LEDGERBAL_HEADER As String
Dim OFX_LEDGERBAL_BALAMT As String
Dim OFX_LEDGERBAL_DTASOF As String
Dim OFX_LEDGERBAL_FOOTER As String

' Closing tag
Dim OFX_FOOTER As String
Dim OFX_STMTRS_FOOTER As String
Dim OFX_STMTTRNRS_FOOTER As String


' OFX file header
OFX_HEADER = "OFXHEADER:100" & Chr(13) & _
                "DATA:OFXSGML" & Chr(13) & _
                "VERSION:102" & Chr(13) & _
                "SECURITY:NONE" & Chr(13) & _
                "ENCODING:USASCII" & Chr(13) & _
                "CHARSET:1252" & Chr(13) & _
                "COMPRESSION:NONE" & Chr(13) & _
                "OLDFILEUID:NONE" & Chr(13) & _
                "NEWFILEUID:NONE" & Chr(13) & _
                "<OFX>" & Chr(13)

OFX_SIGNONMSGSRSV1_HEADER = Chr(9) & "<SIGNONMSGSRSV1>" & Chr(13) & _
                            Chr(9) & Chr(9) & "<SONRS>" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "<STATUS>" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<CODE>0" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<SEVERITY>INFO" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "</STATUS>"
OFX_SIGNONMSGSRSV1_DTSERVER = Chr(9) & Chr(9) & Chr(9) & "<DTSERVER>"
OFX_SIGNONMSGSRSV1_FOOTER = Chr(9) & Chr(9) & Chr(9) & "<LANGUAGE>ENG" & Chr(13) & _
                            Chr(9) & Chr(9) & "</SONRS>" & Chr(13) & _
                            Chr(9) & "</SIGNONMSGSRSV1>"


OFX_BANKMSGSRSV1_HEADER = Chr(9) & "<BANKMSGSRSV1>" & Chr(13) & _
                            Chr(9) & Chr(9) & "<STMTTRNRS>" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "<TRNUID>0" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "<STATUS>" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<CODE>0" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<SEVERITY>INFO" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "</STATUS>" & Chr(13) & _
                            Chr(9) & Chr(9) & "<STMTRS>" & Chr(13) & _
                            Chr(9) & Chr(9) & Chr(9) & "<CURDEF>"
OFX_STMTRS_FOOTER = Chr(9) & Chr(9) & Chr(9) & "</STMTRS>"
OFX_STMTTRNRS_FOOTER = Chr(9) & Chr(9) & Chr(9) & "</STMTTRNRS>"
OFX_BANKMSGSRSV1_FOOTER = Chr(9) & "</BANKMSGSRSV1>"

' Bank account information
OFX_BANKACCTFROM_HEADER = Chr(9) & Chr(9) & Chr(9) & "<BANKACCTFROM>"
OFX_BANKID = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<BANKID>"
OFX_ACCTID = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<ACCTID>"
OFX_ACCTTYPE = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<ACCTTYPE>"
OFX_BANKACCTFROM_FOOTER = Chr(9) & Chr(9) & Chr(9) & "</BANKACCTFROM>"

' Transaction list information
OFX_BANKTRANLIST_HEADER = Chr(9) & Chr(9) & Chr(9) & "<BANKTRANLIST>"
OFX_BANKTRANLIST_DTSTART = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<DTSTART>"
OFX_BANKTRANLIST_DTEND = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<DTEND>"
OFX_BANKTRANLIST_FOOTER = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "</BANKTRANLIST>"

' Transactions information
OFX_STMTTRN_HEADER = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<STMTTRN>"
OFX_STMTTRN_TRNTYPE = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<TRNTYPE>"
OFX_STMTTRN_DTPOSTED = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<DTPOSTED>"
OFX_STMTTRN_TRNAMT = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<TRNAMT>"
OFX_STMTTRN_FITID = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<FITID>"
OFX_STMTTRN_NAME = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<NAME>"
OFX_STMTTRN_MEMO = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<MEMO>"
OFX_STMTTRN_FOOTER = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "</STMTTRN>"

' Ledger balance information
OFX_LEDGERBAL_HEADER = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<LEDGERBAL>"
OFX_LEDGERBAL_BALAMT = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<BALAMT>"
OFX_LEDGERBAL_DTASOF = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "<DTASOF>"
OFX_LEDGERBAL_FOOTER = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "</LEDGERBAL>"

' Closing tag
OFX_FOOTER = "</OFX>"

' Open output file
'OutputFilename = Worksheets("XLS2OFX").Range("OutputFile")
OutputFilename = CurrentWB.Path & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & ".ofx"

Dim fs
Dim ofxstream

Set fs = CreateObject("Scripting.FileSystemObject")
Set ofxstream = fs.CreateTextFile(OutputFilename, True)

Dim iReturn As Integer
Dim StatementStartDate As String
Dim StatementEndDate As String


If Err.Number <> 0 Then
    iReturn = MsgBox(Error(Err.Number), vbCritical, "XLS2OFX Runtime Error")
Else
    ' Write OFX Header
    ofxstream.WriteLine (OFX_HEADER)
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_HEADER)
    'Date format is 20031010000000.000[-5:EST]
    'dtDatetime = Worksheets("XLS2OFX").Range("StatementStartDate")
    dtDatetime = ActiveSheet.Range("A2")
    StatementStartDate = Format(dtDatetime, "yyyymmdd") & "000000.000[-5:EST]"
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_DTSERVER & StatementStartDate)
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_FOOTER)
    
    Dim AcctCurrency As String
    AcctCurrency = "AUD" '= Worksheets("XLS2OFX").Range("AcctCurrency")
    ofxstream.WriteLine (OFX_BANKMSGSRSV1_HEADER & AcctCurrency)
    
    ofxstream.WriteLine (OFX_BANKACCTFROM_HEADER)
    Dim BankId As String
    BankId = "BANK" 'Worksheets("XLS2OFX").Range("BankId")
    ofxstream.WriteLine (OFX_BANKID & BankId)
    Dim AccountNo As String
    AccountNo = "ACCOUNT" 'Worksheets("XLS2OFX").Range("AccountNo")
    ofxstream.WriteLine (OFX_ACCTID & AccountNo)
    Dim AcctType As String
    AcctType = "STATEMENT" 'Worksheets("XLS2OFX").Range("AcctType")
    ofxstream.WriteLine (OFX_ACCTTYPE & AcctType)
    ofxstream.WriteLine (OFX_BANKACCTFROM_FOOTER)
    
    ' Write financial transactions
    ofxstream.WriteLine (OFX_BANKTRANLIST_HEADER)
    ofxstream.WriteLine (OFX_BANKTRANLIST_DTSTART & StatementStartDate)
    'StatementStartDate = Worksheets("XLS2OFX").Range("StatementEndDate")
    dtDatetime = Cells(Rows.Count, 1).End(xlUp).Value 'Worksheets("XLS2OFX").Range("StatementEndDate ")
    StatementEndDate = Format(dtDatetime, "yyyymmdd") & "000000.000[-5:EST]"
    ofxstream.WriteLine (OFX_BANKTRANLIST_DTEND & StatementEndDate)
    
    Dim PreviousBalance
    PreviousBalance = 0 'Worksheets("XLS2OFX").Range("PreviousBalance")
    Dim FinalBalance
    FinalBalance = PreviousBalance
    iTransaction = 1
    
    Dim rgeTransactionList As Range
    Set rgeTransactionList = CurrentWB.ActiveSheet.Range("A1")
    
    'Is Credit Card?
    Dim tfCreditCard As Boolean
        
    tfCreditCard = MsgBox("Is this statement a Credit Card?", vbYesNo)
    
    
    While rgeTransactionList.Offset(iTransaction, 0).Value <> ""
        'Get transaction information
        dtDatetime = rgeTransactionList.Offset(iTransaction, 0).Value ' Worksheets("XLS2OFX").Range("TransactionListTop").Offset(iTransaction, 0).Value
        Dim sTranDate
        sTranDate = Format(dtDatetime, "yyyymmdd") & "000000.000[-5:EST]"
        Dim sTranName As String
        sTranName = rgeTransactionList.Offset(iTransaction, 1).Value
        Dim sTranMemo
        sTranMemo = "" 'rgeTransactionList.Offset(iTransaction, 2).Value
        
        Dim sTranAmount As Double
        sTranAmount = rgeTransactionList.Offset(iTransaction, 2).Value
        
        'Record transaction in OFX format
        ofxstream.WriteLine (OFX_STMTTRN_HEADER)
        
        If tfCreditCard Then
            If sTranAmount > 0 Then
                ofxstream.WriteLine (OFX_STMTTRN_TRNTYPE & "CREDIT")
                sTranAmount = sTranAmount * -1
                FinalBalance = FinalBalance + Val(sTranAmount)
            Else
                ofxstream.WriteLine (OFX_STMTTRN_TRNTYPE & "DEBIT")
                sTranAmount = sTranAmount * -1
                FinalBalance = FinalBalance - Val(sTranAmount)
            End If
        Else
            If sTranAmount < 0 Then
                ofxstream.WriteLine (OFX_STMTTRN_TRNTYPE & "CREDIT")
                sTranAmount = sTranAmount
                FinalBalance = FinalBalance + Val(sTranAmount)
            Else
                ofxstream.WriteLine (OFX_STMTTRN_TRNTYPE & "DEBIT")
                sTranAmount = sTranAmount
                FinalBalance = FinalBalance - Val(sTranAmount)
            End If
        End If
        
        ofxstream.WriteLine (OFX_STMTTRN_DTPOSTED & sTranDate)
        ofxstream.WriteLine (OFX_STMTTRN_TRNAMT & sTranAmount)
        'Emulated FTID format is date stamp plus transaction number (should be unique)
        'Example: "200303170001"
        Dim sTranFTID
        sTranFTID = Format(dtDatetime, "yyyymmdd") & Format(iTransaction, "0000")
        ofxstream.WriteLine (OFX_STMTTRN_FITID & sTranFTID)
        ofxstream.WriteLine (OFX_STMTTRN_NAME & sTranName)
        If Len(sTranMemo) > 0 Then
            ofxstream.WriteLine (OFX_STMTTRN_MEMO & sTranMemo)
        End If
        ofxstream.WriteLine (OFX_STMTTRN_FOOTER)
        
        'Get next transaction
        iTransaction = iTransaction + 1
    Wend
    
    ofxstream.WriteLine (OFX_BANKTRANLIST_FOOTER)
    'Ledger balance
    ofxstream.WriteLine (OFX_LEDGERBAL_HEADER)
    ofxstream.WriteLine (OFX_LEDGERBAL_BALAMT & Str(0))
    ofxstream.WriteLine (OFX_LEDGERBAL_DTASOF & StatementEndDate)
    ofxstream.WriteLine (OFX_LEDGERBAL_FOOTER)
    
    ' Write OFX Footer
    ofxstream.WriteLine (OFX_STMTRS_FOOTER)
    ofxstream.WriteLine (OFX_STMTTRNRS_FOOTER)
    ofxstream.WriteLine (OFX_BANKMSGSRSV1_FOOTER)
    ofxstream.WriteLine (OFX_FOOTER)
    
    ' Close file
    ofxstream.Close
End If

End Sub





