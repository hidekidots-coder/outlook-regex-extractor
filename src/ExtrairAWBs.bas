Option Explicit


' Purpose:
'   Extract AWB / HAWB references from Outlook emails,
'   download attachments and export the references to Excel.
'
' Technologies:
'   VBA
'   Microsoft Outlook Object Model
'   Regular Expressions
'   Microsoft Excel
'
' ============================================================

Public Sub ExtrairAWBs()

    On Error GoTo ErrorHandler

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Const ACCOUNT_NAME As String = "YOUR_OUTLOOK_ACCOUNT"
    Const ROOT_FOLDER As String = "dhl"
    Const FOLDER_1 As String = "notificação de embarque"
    Const FOLDER_2 As String = "pre alerta carga"
    Const DOWNLOAD_FOLDER As String = "C:\Temp\AWB\"
    
    Dim DataInicial As Date
    DataInicial = DateSerial(2026, 8, 15)
    
    ' --------------------------------------------------------
    ' Variables
    ' --------------------------------------------------------
    
    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim RootFolder As Object
    Dim DHLFolder As Object
    Dim Folder As Object
    Dim Email As Object
    Dim Attachment As Object
    
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    
    Dim Ws As Worksheet
    Dim Linha As Long
    Dim TotalEmails As Long
    Dim TotalAWBs As Long
    Dim TotalAttachments As Long
    
    Dim EmailText As String
    Dim AWB As String
    Dim AttachmentPath As String
    
    ' --------------------------------------------------------
    ' Prepare worksheet
    ' --------------------------------------------------------
    
    Set Ws = ActiveSheet
    
    Ws.Columns("A:D").ClearContents
    
    Ws.Cells(1, 1).Value = "AWB / HAWB"
    Ws.Cells(1, 2).Value = "Email Date"
    Ws.Cells(1, 3).Value = "Subject"
    Ws.Cells(1, 4).Value = "Folder"
    
    Linha = 2
    
    ' --------------------------------------------------------
    ' Create download folder
    ' --------------------------------------------------------
    
    CriarPasta DOWNLOAD_FOLDER
    
    ' --------------------------------------------------------
    ' Initialize Outlook
    ' --------------------------------------------------------
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    
    Set RootFolder = OutlookNamespace.Folders(ACCOUNT_NAME)
    
    If RootFolder Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
                  "Outlook account not found: " & ACCOUNT_NAME
    End If
    
    ' --------------------------------------------------------
    ' Find DHL folder
    ' --------------------------------------------------------
    
    Set DHLFolder = BuscarPasta(RootFolder, ROOT_FOLDER)
    
    If DHLFolder Is Nothing Then
        Err.Raise vbObjectError + 1001, , _
                  "Folder not found: " & ROOT_FOLDER
    End If
    
    ' --------------------------------------------------------
    ' Create Regular Expression
    ' --------------------------------------------------------
    
    Set RegEx = CreateAWBRegex()
    
    ' --------------------------------------------------------
    ' Process first folder
    ' --------------------------------------------------------
    
    Set Folder = BuscarPasta(DHLFolder, FOLDER_1)
    
    If Not Folder Is Nothing Then
        
        ProcessarPasta _
            Folder, _
            DataInicial, _
            RegEx, _
            Ws, _
            Linha, _
            TotalEmails, _
            TotalAWBs, _
            TotalAttachments, _
            DOWNLOAD_FOLDER
        
    End If
    
    ' --------------------------------------------------------
    ' Process second folder
    ' --------------------------------------------------------
    
    Set Folder = BuscarPasta(DHLFolder, FOLDER_2)
    
    If Not Folder Is Nothing Then
        
        ProcessarPasta _
            Folder, _
            DataInicial, _
            RegEx, _
            Ws, _
            Linha, _
            TotalEmails, _
            TotalAWBs, _
            TotalAttachments, _
            DOWNLOAD_FOLDER
        
    End If
    
    ' --------------------------------------------------------
    ' Format output
    ' --------------------------------------------------------
    
    Ws.Columns("A:D").AutoFit
    
    Ws.Rows(1).Font.Bold = True
    
    ' --------------------------------------------------------
    ' Finish
    ' --------------------------------------------------------
    
    MsgBox _
        "Process completed successfully!" & vbCrLf & vbCrLf & _
        "Emails processed: " & TotalEmails & vbCrLf & _
        "AWB/HAWB extracted: " & TotalAWBs & vbCrLf & _
        "Attachments downloaded: " & TotalAttachments, _
        vbInformation, _
        "AWB Extractor"

CleanExit:
    Set Match = Nothing
    Set Matches = Nothing
    Set RegEx = Nothing
    
    Set Attachment = Nothing
    Set Email = Nothing
    Set Folder = Nothing
    Set DHLFolder = Nothing
    Set RootFolder = Nothing
    
    Set OutlookNamespace = Nothing
    Set OutlookApp = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox _
        "An error occurred:" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "AWB Extractor Error"
        
    Resume CleanExit
    
End Sub

' ============================================================
' PROCESS OUTLOOK FOLDER
' ============================================================
Private Sub ProcessarPasta( _
    ByVal Folder As Object, _
    ByVal DataInicial As Date, _
    ByVal RegEx As Object, _
    ByVal Ws As Worksheet, _
    ByRef Linha As Long, _
    ByRef TotalEmails As Long, _
    ByRef TotalAWBs As Long, _
    ByRef TotalAttachments As Long, _
    ByVal DownloadFolder As String)

    Dim Email As Object
    Dim Attachment As Object
    
    Dim Matches As Object
    Dim Match As Object
    
    Dim TextoCompleto As String
    Dim AttachmentPath As String
    
    For Each Email In Folder.Items
        
        ' ----------------------------------------------------
        ' Make sure item is an email
        ' Outlook MailItem = Class 43
        ' ----------------------------------------------------
        
        If Email.Class = 43 Then
            
            If Email.ReceivedTime >= DataInicial Then
                
                TotalEmails = TotalEmails + 1
                
                ' ------------------------------------------------
                ' Combine searchable email content
                ' ------------------------------------------------
                
                TextoCompleto = _
                    Email.Subject & " " & _
                    Email.Body & " " & _
                    Email.HTMLBody
                
                ' ------------------------------------------------
                ' Extract AWB / HAWB
                ' ------------------------------------------------
                
                If RegEx.Test(TextoCompleto) Then
                    
                    Set Matches = RegEx.Execute(TextoCompleto)
                    
                    For Each Match In Matches
                        
                        Ws.Cells(Linha, 1).Value = Match.SubMatches(0)
                        Ws.Cells(Linha, 2).Value = Email.ReceivedTime
                        Ws.Cells(Linha, 3).Value = Email.Subject
                        Ws.Cells(Linha, 4).Value = Folder.Name
                        
                        Linha = Linha + 1
                        TotalAWBs = TotalAWBs + 1
                        
                    Next Match
                    
                End If
                
                ' ------------------------------------------------
                ' Download attachments
                ' ------------------------------------------------
                
                If Email.Attachments.Count > 0 Then
                    
                    For Each Attachment In Email.Attachments
                        
                        AttachmentPath = _
                            DownloadFolder & Attachment.FileName
                        
                        Attachment.SaveAsFile AttachmentPath
                        
                        TotalAttachments = TotalAttachments + 1
                        
                    Next Attachment
                    
                End If
                
            End If
            
        End If
        
    Next Email
    
End Sub

' ============================================================
' CREATE AWB REGEX
' ============================================================
Private Function CreateAWBRegex() As Object

    Dim RegEx As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        
        .Pattern = _
            "\bH?AWB[\s:]*""?((\d{10})|(\d{3}-\d{8}))"
        
        .Global = True
        .IgnoreCase = True
        
    End With
    
    Set CreateAWBRegex = RegEx
    
End Function

' ============================================================
' FIND OUTLOOK FOLDER
' ============================================================
Private Function BuscarPasta( _
    ByVal PastaPai As Object, _
    ByVal NomeProcurado As String) As Object

    Dim SubPasta As Object
    
    For Each SubPasta In PastaPai.Folders
        
        If LCase(Trim(SubPasta.Name)) = _
           LCase(Trim(NomeProcurado)) Then
            
            Set BuscarPasta = SubPasta
            Exit Function
            
        End If
        
    Next SubPasta
    
End Function

' ============================================================
' CREATE FOLDER
' ============================================================
Private Sub CriarPasta(ByVal Caminho As String)

    If Dir(Caminho, vbDirectory) = "" Then
        
        MkDir Caminho
        
    End If
    
End Sub
