Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnSelectSong_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSelectSong.Click
        Try
            Dim getLyrics As New GetLyrics()
            Dim result As String = getLyrics.FileDlg()

            ' Handle the result
            If result IsNot Nothing Then
                ' The result contains a status message from AddLyricsToSlide
                ' It could be "Lyrics added to the slides successfully." or "No 'lyrics' layout found in the slide master."
                MessageBox.Show(result, "Lyrics Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' User cancelled the file dialog
                MessageBox.Show("File selection cancelled.", "Lyrics Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            ' Handle any errors that might occur
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
