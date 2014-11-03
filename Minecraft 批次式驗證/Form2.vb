Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        For k As Integer = 0 To Form1.share.GetLength(0) - 1
            ListBox1.Items.Add(Form1.share(k, 0) & Form1.share(k, 1))
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim num As String
        TextBox8.Clear()
        If CheckBox1.Checked Then
            TextBox8.Text += TextBox6.Text
        End If
        For k As Integer = 0 To Form1.share.GetLength(0) - 1
            num = k + 1
            If CheckBox2.Checked Then
                If CheckBox2.Checked And CheckBox3.Checked Then
                    num = num.ToString.PadLeft(NumericUpDown1.Value, "0")
                End If
                If CheckBox2.Checked And CheckBox4.Checked Then

                End If
                Dim outl As String
                outl = TextBox1.Text
                outl = outl.Replace("%S", num)
                outl = outl.Replace("%N", Form1.share(k, 0))
                outl = outl.Replace("%R", Form1.share(k, 1))
                If CheckBox4.Checked Then
                    outl = outl.Replace("不是正版", TextBox3.Text)
                    outl = outl.Replace("是正版", TextBox2.Text)
                    outl = outl.Replace("測試失敗，原因：網路錯誤", TextBox4.Text)
                    outl = outl.Replace("測試失敗，原因：其他錯誤", TextBox5.Text)
                End If
                TextBox8.Text += outl & Chr(13) & Chr(10)
            Else
                TextBox8.Text += k + 1 & ". " & Form1.share(k, 0) & Form1.share(k, 1) & Chr(13) & Chr(10)
            End If
        Next
        If CheckBox5.Checked Then
            TextBox8.Text += TextBox7.Text
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBox6.Enabled = True
        Else
            TextBox6.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            TextBox1.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
        Else
            TextBox1.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            TextBox7.Enabled = True
        Else
            TextBox7.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
        Else
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            NumericUpDown1.Enabled = True
        Else
            NumericUpDown1.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Clipboard.Clear()
        Clipboard.SetText(TextBox8.Text)
        MsgBox("已複製結果", MsgBoxStyle.Information)
    End Sub
    Private Sub ChartUpadta() Handles TrackBar1.Scroll, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, TextBox9.TextChanged, CheckBox6.CheckedChanged, MyBase.Shown
        Dim out As String = _
         "<html>" & Chr(10) & _
                "<head>" & Chr(10) & _
                    "<META http-equiv=""Content-Type"" content=""text/html; charset=UTF8"">" & Chr(10) & _
                    "<script type=""text/javascript"" src=""https://www.google.com/jsapi""></script>" & Chr(10) & _
                    "<script type=""text/javascript"">" & Chr(10) & _
                        "google.load(""visualization"", ""1"", {packages:[""corechart""]});" & Chr(10) & _
                        "google.setOnLoadCallback(drawChart);" & Chr(10) & _
                        "function drawChart() {" & Chr(10) & _
                            "var data = google.visualization.arrayToDataTable([" & Chr(10) & _
                                "['項目', '人數']," & Chr(10) & _
                                "['%tp人是正版',%tp]," & Chr(10) & _
                                "['%tnp人是盜版',  %tnp]," & Chr(10) & _
                                "['%fn人測試失敗：網路錯誤',%fn]," & Chr(10) & _
                                "['%fo人測試失敗：其他錯誤', %fo]" & Chr(10) & _
                            "]);" & Chr(10) & _
                            "var options = {" & Chr(10) & _
                                "title:  '%Title'," & Chr(10) & _
                                "is3D: %3D," & Chr(10) & _
                                "pieHole: '%Hole'" & Chr(10) & _
                            "};" & Chr(10) & _
                            "var chart = new google.visualization.PieChart(document.getElementById('piechart'));" & Chr(10) & _
                            "chart.draw(data, options);" & Chr(10) & _
                        "}" & Chr(10) & _
                    "</script>" & Chr(10) & _
                "</head>" & Chr(10) & _
                "<body>" & Chr(10) & _
                    "<div id=""piechart"" style=""width: 490px; height: 330px;""></div>" & Chr(10) & _
                "</body>" & Chr(10) & _
            "</html>​"
        out = out.Replace("%tp", Form1.tp)
        out = out.Replace("%tnp", Form1.tnp)
        out = out.Replace("%fn", Form1.fn)
        out = out.Replace("%fo", Form1.fo)
        out = out.Replace("%Title", TextBox9.Text)
        out = out.Replace("%3D", RadioButton2.Checked.ToString.ToLower)
        out = out.Replace("%Hole", TrackBar1.Value / 100)
        WebBrowser1.DocumentText = out
    End Sub
    Private Sub MoveLabel(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label9.Text = Format(TrackBar1.Value / 100, "0%")
        Label9.Location = New Point((TrackBar1.Value * 4.79) + 13, 144)
    End Sub
    Private Sub Tracke(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged, RadioButton1.CheckedChanged
        If CheckBox6.Checked And RadioButton1.Checked Then
            TrackBar1.Enabled = True
        Else
            TrackBar1.Enabled = False
        End If

        If RadioButton1.Checked Then
            CheckBox6.Enabled = True
        Else
            CheckBox6.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If WebBrowser1.ReadyState = WebBrowserReadyState.Complete Then

            Dim Height As Integer = WebBrowser1.Document.Body.ScrollRectangle.Height
            Dim Width As Integer = WebBrowser1.Document.Body.ScrollRectangle.Width

            WebBrowser1.Height = Height
            WebBrowser1.Width = Width

            Dim Bitmap As Bitmap = New Bitmap(Width, Height)
            Dim Rectangle As Rectangle = New Rectangle(0, 0, Width, Height)
            WebBrowser1.DrawToBitmap(Bitmap, Rectangle)

            Dim SaveFileDialog1 As SaveFileDialog = New SaveFileDialog()
            SaveFileDialog1.Filter = "JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png"
            SaveFileDialog1.ShowDialog()

            Bitmap.Save(SaveFileDialog1.FileName)
        End If
    End Sub

    Function Base64ToImage(ByVal base64string As String) As System.Drawing.Image
        'Setup image and get data stream together
        Dim img As System.Drawing.Image
        Dim MS As System.IO.MemoryStream = New System.IO.MemoryStream
        Dim b64 As String = base64string.Replace(" ", "+")
        Dim b() As Byte

        'Converts the base64 encoded msg to image data
        b = Convert.FromBase64String(b64)
        MS = New System.IO.MemoryStream(b)

        'creates image
        img = System.Drawing.Image.FromStream(MS)

        Return img
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form3.Show()
    End Sub
End Class