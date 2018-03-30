'-----------------------------------------------------------------------------------------------------------------------------------------------
'
'	This program is free software; you can redistribute it and/or
'	modify it under the terms of the GNU General Public License
'	as published by the Free Software Foundation; either version 2
'	of the License, or (at your option) any later version.
'
'	This program is distributed in the hope that it will be useful,
'	but WITHOUT ANY WARRANTY; without even the implied warranty of
'	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	GNU General Public License for more details.
'
'	You should have received a copy of the GNU General Public License
'	along with this program; if not, write to the Free Software Foundation,
'	Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.
'
'	Name :
'				BEE : BatchEmailsExporter
'	Author :
'				▄▄▄▄▄▄▄  ▄ ▄▄ ▄▄▄▄▄▄▄
'				█ ▄▄▄ █ ██ ▀▄ █ ▄▄▄ █
'				█ ███ █ ▄▀ ▀▄ █ ███ █
'				█▄▄▄▄▄█ █ ▄▀█ █▄▄▄▄▄█
'				▄▄ ▄  ▄▄▀██▀▀ ▄▄▄ ▄▄
'				 ▀█▄█▄▄▄█▀▀ ▄▄▀█ █▄▀█
'				 █ █▀▄▄▄▀██▀▄ █▄▄█ ▀█
'				▄▄▄▄▄▄▄ █▄█▀ ▄ ██ ▄█
'				█ ▄▄▄ █  █▀█▀ ▄▀▀  ▄▀
'				█ ███ █ ▀▄  ▄▀▀▄▄▀█▀█
'				█▄▄▄▄▄█ ███▀▄▀ ▀██ ▄
'
'-----------------------------------------------------------------------------------------------------------------------------------------------

Imports System.IO
Imports Microsoft.Office.Interop

Public Class MainForm

    'Main Form loading event
    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ttHoverInfo.SetToolTip(rbDoc, "Microsoft Office Word format (.doc)")
        ttHoverInfo.SetToolTip(rbHtml, "HTML format (.htm)")
        ttHoverInfo.SetToolTip(rbMht, "MIME HTML format (.mht)")
        ttHoverInfo.SetToolTip(rbMsg, "Outlook Unicode message format (.msg)")
        ttHoverInfo.SetToolTip(rbRtf, "Rich Text format (.rtf)")
        ttHoverInfo.SetToolTip(rbTxt, "Text format (.txt)")
        ttHoverInfo.SetToolTip(cbStop, "Stop emails export")
        ttHoverInfo.SetToolTip(cbStart, "Start emails export")
        ttHoverInfo.SetToolTip(cbFolderSelect, "Select output folder")
        ttHoverInfo.SetToolTip(cbAttachments, "Save emails attachments")
        ttHoverInfo.SetToolTip(cbAbout, "About " & My.Application.Info.AssemblyName)
        ttHoverInfo.SetToolTip(tbFolderPath, "Output folder path")
    End Sub

    'If "Stop" button is clicked :
    ' - stop emails export
    ' - unlock main Form controls
    Private Sub cbStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStop.Click
        bStopThread = True
    End Sub

    'If "?" button is clicked, show the "About" Form
    Private Sub cbAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAbout.Click
        ABOUT.Show()
    End Sub

    'If "Start" button is clicked :
    ' - check if an output folder is selected
    ' - check if the output folder exists
    ' - lock main Form controls
    ' - export emails from selected Outlook folder
    Private Sub cbStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStart.Click

        Dim sSelectedFolderPath As String = tbFolderPath.Text
        If sSelectedFolderPath = String.Empty Then
            MsgBox("Please select an output folder.", vbInformation, "Invalid output folder")
            Exit Sub
        End If
        If Not System.IO.Directory.Exists(sSelectedFolderPath) Then
            MsgBox("Output folder path does not exist.", vbCritical, "Invalid output folder")
            Exit Sub
        End If

        Call lockUI()

        Dim sExtension As String
        Dim oSaveFormat As Outlook.OlSaveAsType
        If rbHtml.Checked Then
            sExtension = "htm"
            oSaveFormat = Outlook.OlSaveAsType.olHTML
        ElseIf rbMht.Checked Then
            sExtension = "mht"
            oSaveFormat = Outlook.OlSaveAsType.olMHTML
        ElseIf rbMsg.Checked Then
            sExtension = "msg"
            oSaveFormat = Outlook.OlSaveAsType.olMSGUnicode
        ElseIf rbDoc.Checked Then
            sExtension = "doc"
            oSaveFormat = Outlook.OlSaveAsType.olDoc
        ElseIf rbRtf.Checked Then
            sExtension = "rtf"
            oSaveFormat = Outlook.OlSaveAsType.olRTF
        Else
            sExtension = "txt"
            oSaveFormat = Outlook.OlSaveAsType.olTXT
        End If

        t = New Threading.Thread(Sub() exportEmails(sSelectedFolderPath, sExtension, oSaveFormat))
        t.Start()

    End Sub

    'If "Output folder" button is clicked, show a folder selection dialog
    Private Sub cbFolderSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFolderSelect.Click
        Dim sSelectedFolderPath As String
        sSelectedFolderPath = getFolderFromDialog()
        If (sSelectedFolderPath <> String.Empty) Then
            tbFolderPath.Text = CStr(sSelectedFolderPath)
        End If
    End Sub

    'Lock main Form controls
    Private Sub lockUI()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf lockUI))
        Else
            progressBarCompletion.Value = progressBarCompletion.Minimum
            rbHtml.Enabled = Not rbHtml.Enabled
            rbMht.Enabled = Not rbMht.Enabled
            rbMsg.Enabled = Not rbMsg.Enabled
            rbDoc.Enabled = Not rbDoc.Enabled
            rbTxt.Enabled = Not rbTxt.Enabled
            rbRtf.Enabled = Not rbRtf.Enabled
            If Not rbMsg.Checked Then
                cbAttachments.Enabled = Not cbAttachments.Enabled
            End If
            tbFolderPath.Enabled = Not tbFolderPath.Enabled
            cbFolderSelect.Enabled = Not cbFolderSelect.Enabled
            cbStart.Enabled = Not cbStart.Enabled
            cbStop.Enabled = Not cbStop.Enabled
        End If
    End Sub

    'Unlock main Form controls
    Private Sub unlockUI()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf unlockUI))
        Else
            progressBarCompletion.Value = progressBarCompletion.Maximum
            rbHtml.Enabled = Not rbHtml.Enabled
            rbMht.Enabled = Not rbMht.Enabled
            rbMsg.Enabled = Not rbMsg.Enabled
            rbDoc.Enabled = Not rbDoc.Enabled
            rbTxt.Enabled = Not rbTxt.Enabled
            rbRtf.Enabled = Not rbRtf.Enabled
            If Not rbMsg.Checked Then
                cbAttachments.Enabled = Not cbAttachments.Enabled
            End If
            tbFolderPath.Enabled = Not tbFolderPath.Enabled
            cbFolderSelect.Enabled = Not cbFolderSelect.Enabled
            cbStart.Enabled = Not cbStart.Enabled
            cbStop.Enabled = Not cbStop.Enabled
        End If
    End Sub

    'Update main Form ProgressBar from thread
    Public Sub updateUI()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf updateUI))
        Else
            progressBarCompletion.Value = iPercent
            progressBarCompletion.Refresh()
        End If
    End Sub

    'Export emails from selected Outlook folder to a given path with a given file format : htm, mht, msg, doc, rtf, txt
    Sub exportEmails(ByVal sSelectedFolderPath As String, ByVal sExtension As String, ByVal oSaveFormat As Outlook.OlSaveAsType)

        iPercent = 0
        Dim iPercentCopy As Integer

        Dim appOutlook As New Outlook.Application
        Dim nsMAPI As Outlook.NameSpace = appOutlook.GetNamespace("MAPI")

        'Get Outlook email folder or account selection
        Dim oOutlookFolder As Outlook.MAPIFolder = nsMAPI.PickFolder
        If oOutlookFolder Is Nothing Then
            MsgBox("Please select an Outlook emails folder.", vbInformation, "Invalid emails folder")
            GoTo exitSub
        End If

        Dim colFolders As New Collection
        Dim colEntryID As New Collection
        Dim colStoreID As New Collection
        Call getOutlookFolder(colFolders, colEntryID, colStoreID, oOutlookFolder)

        Dim sFolderPath As String
        Dim sOutputPath As String
        Dim sCleanFolderPath() As String = New String() {}

        Dim mItem As Outlook.MailItem
        Dim subFolder As Outlook.MAPIFolder
        Dim sEmailName As String
        Dim sReceivedDate As String
        Dim sEmailFilePath As String
        Dim sAttachmentsPath As String

        'For each folder...
        For lFoldersCpt As Long = 1 To colFolders.Count

            sFolderPath = CStr(colFolders(lFoldersCpt))
            sFolderPath = Mid(sFolderPath, InStr(3, sFolderPath, Path.DirectorySeparatorChar) + 1)
            'Ignore empty root account folder
            If Strings.Left(sFolderPath, 2) = "\\" Then
                Continue For
            End If
            'Delete useless spaces and illegal chars from path
            sCleanFolderPath = Split(sFolderPath, Path.DirectorySeparatorChar)
            For i As Integer = 0 To sCleanFolderPath.Length - 1
                sCleanFolderPath(i) = delAllSpace(stripIllegalChars(sCleanFolderPath(i)))
            Next
            sFolderPath = String.Join(Path.DirectorySeparatorChar, sCleanFolderPath)
            sOutputPath = sSelectedFolderPath & Path.DirectorySeparatorChar & sFolderPath

            subFolder = appOutlook.Session.GetFolderFromID(CStr(colEntryID(lFoldersCpt)), colStoreID(lFoldersCpt))

            'Ignore folders not related to emails : contacts, calendars, tasks, notes...
            If Not subFolder.DefaultItemType.ToString = Outlook.OlItemType.olMailItem.ToString Then
                Continue For
            ElseIf Not System.IO.Directory.Exists(sOutputPath) Then
                'Create folder if it does not exist
                System.IO.Directory.CreateDirectory(sOutputPath)
            End If

            iPercentCopy = iPercent

            'For each email in folder...
            For lMailsCpt As Long = 1 To subFolder.Items.Count
                'Ignore objects that are not emails
                If Not TypeOf subFolder.Items(lMailsCpt) Is Outlook.MailItem Then
                    Continue For
                End If
                mItem = subFolder.Items(lMailsCpt)
                sReceivedDate = Format(mItem.ReceivedTime, "yyyyMMdd-hhmmss")
                sEmailName = stripIllegalChars(mItem.Subject)
                'M$ Windows is SHIT, so paths have a 260 chars limit
                'If no attachments and not a html export, let's use 250 chars for file paths
                'If attachments, we store them in a folder with same name as related email file : 200 chars for folder path and 50 chars for filenames (200+50=250)
                If (((Not rbHtml.Checked) And (Not cbAttachments.Checked)) Or (rbMsg.Checked)) Then
                    sEmailFilePath = Strings.Left(delAllSpace(sOutputPath & Path.DirectorySeparatorChar & sReceivedDate & "_" & sEmailName), 250) & "." & sExtension
                Else
                    sEmailFilePath = Strings.Left(delAllSpace(sOutputPath & Path.DirectorySeparatorChar & sReceivedDate & "_" & sEmailName), 200) & "." & sExtension
                End If
                Try
                    'Export email file
                    mItem.SaveAs(sEmailFilePath, oSaveFormat)
                    'Export email attachments
                    If cbAttachments.Checked And Not rbMsg.Checked Then
                        If mItem.Attachments.Count > 0 Then
                            For j As Long = 1 To mItem.Attachments.Count
                                'A folder name can't end with a space or a dot
                                sAttachmentsPath = Strings.Left(delAllSpace(sOutputPath & Path.DirectorySeparatorChar & sReceivedDate & "_" & sEmailName), 200)
                                If Not System.IO.Directory.Exists(sAttachmentsPath) Then
                                    System.IO.Directory.CreateDirectory(sAttachmentsPath)
                                End If
                                mItem.Attachments(j).SaveAsFile(Strings.Left(sAttachmentsPath & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(mItem.Attachments(j).FileName), 250) & Path.GetExtension(mItem.Attachments(j).FileName))
                            Next
                        End If
                    End If
                Catch ex As System.Exception
                    'Oops
                End Try

                'More precise completion percentage
                iPercent = iPercentCopy + CInt((lMailsCpt * (((lFoldersCpt + 1) * 100 / colFolders.Count) - (lFoldersCpt * 100 / colFolders.Count)) / subFolder.Items.Count))
                Call updateUI()

                'Abort, "Stop" button has been clicked
                If bStopThread Then
                    GoTo exitSub
                End If
            Next (lMailsCpt)

            'Completion percentage
            iPercent = CInt(lFoldersCpt * 100 / colFolders.Count)
            Call updateUI()

            'Abort, "Stop" button has been clicked
            If bStopThread Then
                GoTo exitSub
            End If

        Next lFoldersCpt

        'Easter egg
        My.Computer.Audio.Play(My.Resources.complete, AudioPlayMode.Background)
        MsgBox("Export completed successfully !", MsgBoxStyle.SystemModal And vbInformation, "Done")

exitSub:
        bStopThread = False
        Call unlockUI()

    End Sub

    'If "msg" is checked, disabled the "Save attachments" checkbox because attachments are already included in .msg unicode file format
    Private Sub rb_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMsg.CheckedChanged, rbMht.CheckedChanged, rbHtml.CheckedChanged, rbTxt.CheckedChanged, rbDoc.CheckedChanged, rbRtf.CheckedChanged
        If rbMsg Is sender Then
            cbAttachments.Enabled = False
        Else
            cbAttachments.Enabled = True
        End If
    End Sub

End Class