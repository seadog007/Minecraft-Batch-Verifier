Imports System.Net
Imports System.IO
Imports System.Text
Public Class Form1
    Public share(,) As String
    Public tp As Integer = 0
    Public tnp As Integer = 0
    Public fn As Integer = 0
    Public fo As Integer = 0
    Public all As Integer = 0
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim sr As New System.IO.StreamReader(OpenFileDialog1.FileName)
            TextBox1.Text = sr.ReadToEnd
            sr.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("名單不可為空", MsgBoxStyle.Critical)
        Else
            tp = 0
            tnp = 0
            fn = 0
            fo = 0
            all = 0
            TextBox2.Clear()
            Dim Test() As String = Split(TextBox1.Text.Replace(Chr(10), ""), Chr(13))
            ReDim share(Test.Length - 1, 1)
            all = Test.Length
            For i As Integer = 0 To Test.Length - 1
                Dim Thread As New System.Threading.Thread(AddressOf Log)
                Thread.Start(i + 1.ToString & "〰" & Test(i))
                If i = Test.Length - 1 Then
                    Thread.Join()
                End If
            Next
            MsgBox("將啟動自動排列系統", MsgBoxStyle.Information)
            Resort()
            If Not ((tp + tnp + fn + fo) = all) Then
                MsgBox("重新統計人數中", MsgBoxStyle.Information)
            Else
                MsgBox("分析結果" & Chr(13) & Chr(10) & "正版玩家共" & tp & "人" & pr("tp") & Chr(13) & Chr(10) & "盜版玩家共" & tnp & "人" & pr("tnp") & Chr(13) & Chr(10) & "因網路錯誤而無法完成分析者有" & fn & "人" & pr("fn") & Chr(13) & Chr(10) & "因其他錯誤而無法完成分析者有" & fo & "人" & pr("fo") & Chr(13) & Chr(10) & "總計" & all & "人", MsgBoxStyle.Information)
            End If
            Button4.Enabled = True
        End If
    End Sub
    Public Sub Resort()
        TextBox2.Clear()
        For k As Integer = 0 To share.GetLength(0) - 1
            TextBox2.Text += k + 1 & ". " & share(k, 0) & share(k, 1) & Chr(13) & Chr(10)
        Next
    End Sub
    Public Sub Log(ByVal ID As String)
        Dim a() As String = Split(ID, "〰")
        Dim num As Integer = a(0)
        ID = a(1)
        Dim re As String = checkplus(ID)
        TextBox2.Text += num & ". " & ID & re & Chr(13) & Chr(10)
        share(num - 1, 0) = ID
        share(num - 1, 1) = re
    End Sub
    Public Function checkplus(ByVal ID As String)
        Dim checkurl As String = "http://www.minecraft.net/haspaid.jsp?user=" & ID
        Dim StockValues As String = SketchWebPage(checkurl)
        If StockValues = "FAILED" Then
            fn += 1
            Return "測試失敗，原因：網路錯誤"
        ElseIf StockValues = "true" Then
            tp += 1
            Return "是正版"
        ElseIf StockValues = "false" Then
            tnp += 1
            Return "不是正版"
        Else
            fo += 1
            Return "測試失敗，原因：其他錯誤"
        End If
    End Function
    Private Function SketchWebPage(ByVal URL As String) As String
        Try
            Dim lobjRequest As HttpWebRequest
            Dim lobjResponse As HttpWebResponse
            Dim lobjEncode As Encoding
            Dim lobjStreamReader As StreamReader
            lobjRequest = CType(WebRequest.Create(URL), HttpWebRequest)
            lobjResponse = CType(lobjRequest.GetResponse(), HttpWebResponse)
            lobjEncode = System.Text.Encoding.GetEncoding("big5")
            '建立一個新的stream去做讀取
            lobjStreamReader = New StreamReader(lobjResponse.GetResponseStream, lobjEncode)
            Dim stmPage As String = lobjStreamReader.ReadToEnd()
            lobjResponse.Close()
            lobjStreamReader.Close()
            Return stmPage
        Catch ex As Exception
            Return "FAILED"
        End Try
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form.CheckForIllegalCrossThreadCalls = False
        Dim Thread1 As New System.Threading.Thread(AddressOf GetInfo_Post)
        Thread1.Start()
    End Sub
    Public Function pr(ByVal o As String)
        If o = "tp" Then
            Return Format(tp / all * 100, "#0.000") & "%"
        ElseIf o = "tnp" Then
            Return Format(tnp / all * 100, "#0.000") & "%"
        ElseIf o = "fn" Then
            Return Format(fn / all * 100, "#0.000") & "%"
        ElseIf o = "fo" Then
            Return Format(fo / all * 100, "#0.000") & "%"
        Else
            Return "N/A"
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Clipboard.Clear()
        Clipboard.SetText(TextBox2.Text)
        MsgBox("已複製結果", MsgBoxStyle.Information)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.ShowDialog()
    End Sub
End Class
