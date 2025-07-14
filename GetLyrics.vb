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

        Dim result As String = ProcessLyrics(lyrics)
        Return result
    End Function

    Public Function PasteDlg()

        ' open textarea dialog
        Dim inputForm As New Form()
        Dim textBox As New TextBox With {
            .Multiline = True,
            .Dock = DockStyle.Fill,
            .ScrollBars = ScrollBars.Vertical
        }

        Dim lyrics As String = textBox.Text
        Dim result As String = ProcessLyrics(lyrics)

        inputForm.Controls.Add(textBox)
        inputForm.Text = "Paste Lyrics"
        inputForm.Width = 600
        inputForm.Height = 400
        Dim btnOk As New Button With {
            .Text = "OK",
            .Dock = DockStyle.Bottom
        }
        inputForm.Controls.Add(btnOk)
        AddHandler btnOk.Click, Sub(sender, e)
                                    inputForm.DialogResult = DialogResult.OK
                                    inputForm.Close()

                                End Sub
        Dim btnCancel As New Button With {
            .Text = "Cancel",
            .Dock = DockStyle.Bottom
        }
        inputForm.Controls.Add(btnCancel)
        AddHandler btnCancel.Click, Sub(sender, e)
                                        inputForm.DialogResult = DialogResult.Cancel
                                        inputForm.Close()
                                    End Sub

        'if user clicks OK, return the lyrics
        If inputForm.ShowDialog() = DialogResult.OK Then
            lyrics = textBox.Text
            Return ProcessLyrics(lyrics)
        Else
            Return Nothing
        End If

    End Function

    Public Function ProcessLyrics(lyrics As String) As String
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
