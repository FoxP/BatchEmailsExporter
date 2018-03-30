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

Module SharedCode
    Public iPercent As Integer
    Public t As Threading.Thread
    Public bStopThread As Boolean

    'A better alternative to FolderBrowserDialog, using a OpenFileDialog in a "hacky" way
    Function getFolderFromDialog(Optional ByVal sDialogTitle As String = "Select a folder", _
             Optional ByVal sInitialDirectory As String = "::{450d8fba-ad25-11d0-98a8-0800361b1103}", _
             Optional ByVal bRestoreDirectory As Boolean = True _
            ) As String
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = sDialogTitle
        fd.InitialDirectory = sInitialDirectory
        fd.RestoreDirectory = bRestoreDirectory
        fd.FileName = sDialogTitle
        fd.Filter = "Folder|*.folder"
        fd.CheckPathExists = True
        fd.ShowReadOnly = False
        fd.ReadOnlyChecked = True
        fd.CheckFileExists = False
        fd.ValidateNames = False

        If fd.ShowDialog() = DialogResult.OK Then
            Return Path.GetDirectoryName(fd.FileName)
        Else
            Return String.Empty
        End If
    End Function

    'Strip common Microsoft Windows paths illegal chars
    Function stripIllegalChars(ByVal sInput As String) As String
        Dim regexObj As Object
        regexObj = CreateObject("vbscript.regexp")
        regexObj.Pattern = "[\" & Chr(34) & "\>\<\:\.\/\|\?\*\\]"
        regexObj.IgnoreCase = True
        regexObj.Global = True
        Return CStr(regexObj.Replace(sInput, ""))
exitFunction:
        regexObj = Nothing
    End Function

    'Dialog to select and browse an Outlook emails folder
    Sub getOutlookFolder(ByVal colFolders As Collection, ByVal colEntryID As Collection, ByVal colStoreID As Collection, ByVal fld As Outlook.MAPIFolder)
        Dim subFolder As Outlook.MAPIFolder
        colFolders.Add(fld.FolderPath)
        colEntryID.Add(fld.EntryID)
        colStoreID.Add(fld.StoreID)
        For Each subFolder In fld.Folders
            getOutlookFolder(colFolders, colEntryID, colStoreID, subFolder)
        Next subFolder
exitSub:
        subFolder = Nothing
    End Sub

    'Delete useless space characters from a given string
    Function delAllSpace(ByVal strParamString As String) As String

        Dim strTempString As String
        Dim i As Integer

        strTempString = LTrim(strParamString)
        strTempString = RTrim(strTempString)

        i = InStr(1, strTempString, "  ")

        While i <> 0
            strTempString = Replace(strTempString, "  ", " ")
            i = InStr(1, strTempString, "  ")
            System.Windows.Forms.Application.DoEvents()
        End While

        delAllSpace = strTempString

    End Function

End Module
