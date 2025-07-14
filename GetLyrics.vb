Imports System.Windows.Forms
Public Class GetLyrics
    Public Function FileDlg()
        Using dlg As New OpenFileDialog
            dlg.Title = "Select a song to insert."
            dlg.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"

            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

            If dlg.ShowDialog() = DialogResult.OK Then

                'Call ExtractLyrics(selectedFile)
                Return ExtractLyrics(dlg.FileName)

            Else
                Return Nothing
            End If
        End Using
    End Function

    Public Function ExtractLyrics(selectedFile As String)

        Dim lyrics As String = System.IO.File.ReadAllText(selectedFile)
        Dim blocks As New List(Of String)

        ' Split by two newlines (paragraphs)
        Dim parts() As String = lyrics.Split(New String() {Environment.NewLine & Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

        Console.WriteLine("Extracted Lyrics:")

        For Each line As String In parts
            Dim trimmed As String = line.Trim()
            If Not String.IsNullOrWhiteSpace(trimmed) Then
                blocks.Add(trimmed)
                Console.WriteLine(trimmed)
            End If
        Next

        ' call AddToSlide
        Dim addToSlide As New AddToSlide()
        Dim result As String = addToSlide.AddLyricsToSlide(blocks)
        Return result
    End Function

End Class
