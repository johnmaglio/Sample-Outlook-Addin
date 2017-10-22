Imports System.Diagnostics
Imports System.Drawing
Imports stdole




Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        'register function to run when folder context menu is displayed "ClientMatterContextMenuHandler
        AddHandler Me.Application.FolderContextMenuDisplay, AddressOf ClientMatterContextMenuHandler
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub ClientMatterContextMenuHandler(ByVal commandBar As Office.CommandBar, ByVal currentFolder As Outlook.MAPIFolder)
        Dim currentFolderId As String = currentFolder.EntryID
        Try
            If isValidFolder(currentFolderId) Then
                Dim ClientMatterString As String = GetClientMatterForFolder(currentFolderId)
                Dim contextMenuWebDisplay As Office.CommandBarButton = commandBar.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, ClientMatterString, Type.Missing, True)
                contextMenuWebDisplay.Visible = True
                contextMenuWebDisplay.Picture = ImageConverter.ImageToPictureDisp(My.Resources.ClientMatterFolderImage)
                contextMenuWebDisplay.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption
                contextMenuWebDisplay.Caption = "View emails for this matter"
                'register event that will run when the menu button is clicked
                AddHandler contextMenuWebDisplay.Click, AddressOf LaunchIEWithWebPage
            End If
        Catch ex As Exception
            'log exception to database/file
        End Try
    End Sub

    Private Function isValidFolder(folderId As String) As Boolean
        'put logic here to make sure the context menu should be displayed for this folder 
        Return True
    End Function

    Private Function GetClientMatterForFolder(ByVal folderID As String) As String
        'put logic here to determine the client/matter for the given Exchange folder
        'return client/matter separated by semicolon
        Return "1000;1"
    End Function

    Public Sub LaunchIEWithWebPage(ByVal ctrl As Office.CommandBarButton, ByRef cancel As Boolean)
        Dim ClientMatterString As String = ctrl.Parameter
        Dim url As String = "http://www.google.com" 'change link to base search system
        Dim Client As Integer = CInt(ClientMatterString.Split(";")(0))
        Dim Matter As Integer = CInt(ClientMatterString.Split(";")(1))
        If Client <> 0 Then
            'set client parameter in search system
            url += "?clientparameter=" + Client.ToString
            If Matter <> 0 Then
                'set matter parameter in search system
                url += "&matterparameter=" + Matter.ToString
            End If
            'launch IE with url
            Process.Start("IEXPLORE.EXE", "-nomerge " + url)
        End If
    End Sub

    Public Class ImageConverter
        Inherits System.Windows.Forms.AxHost
        Public Sub New()
            MyBase.New("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
        End Sub

        Public Shared Function ImageToPictureDisp(image As image) As ipicturedisp
            Return DirectCast(GetIPictureDispFromPicture(image), IPictureDisp)
        End Function
    End Class

End Class
