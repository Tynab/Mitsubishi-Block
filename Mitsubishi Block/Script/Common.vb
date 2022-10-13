Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Math
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports System.Windows.Forms.Application
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons
Imports System.Windows.Forms.MessageBoxIcon

Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Check internet connection.
    ''' </summary>
    ''' <returns>Connection state.</returns>
    Private Function IsNetAvail()
        Dim objResp As WebResponse
        Try
            objResp = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse
            objResp.Close()
            objResp = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check update.
    ''' </summary>
    Private Sub ChkUpd()
        If IsNetAvail() AndAlso Not (New WebClient).DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
            Show($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", "更新", OK, Information)
            Run(New FrmUpdate)
        End If
    End Sub

    ''' <summary>
    ''' Update valid license
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Fade in form
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity += 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Fade out form
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity -= 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' End process.
    ''' </summary>
    ''' <param name="name">Process name.</param>
    Private Sub KillPrcs(name As String)
        If GetProcessesByName(name).Count > 0 Then
            For Each item In GetProcessesByName(name)
                item.Kill()
            Next
        End If
    End Sub

    ''' <summary>
    ''' Run application.
    ''' </summary>
    Friend Sub RunApp()
        ' Input w
        Dim w = 0D
        For i = 0 To Integer.MaxValue
            Dim wi = HdrDInp(vbTab & $"w{i + 1} = ")
            If Not wi > 0 Then
                If i = 0 Then
                    i = -1
                Else
                    Exit For
                End If
            Else
                w += ConvertToG(wi)
            End If
        Next
        ' Input h
        Dim h = 0D
        For i = 0 To Integer.MaxValue
            Dim hi = HdrDInp(vbTab & $"h{i + 1} = ")
            If Not hi > 0 Then
                If i = 0 Then
                    i = -1
                Else
                    Exit For
                End If
            Else
                h += ConvertToG(hi)
            End If
        Next
        ' Calculate
        Dim c = Ceiling((w + h) * 2 / Pow(10, 3))
        Dim s = Ceiling(w * h / Pow(10, 6))
        Dim block = c + s
        For i = 10 To Integer.MaxValue
            If (block + i) Mod 30 = 0 Then
                block += i
                Exit For
            End If
        Next
        ' Print
        Intro()
        ForegroundColor = DarkCyan
        Dim fmt = FmtNo(c, s, block)
        WriteLine(vbTab & "Ｃ (m)" & vbTab & vbTab & ": " + String.Format(fmt, c))
        WriteLine(vbTab & "Ｓ (m²)" & vbTab & vbTab & ": " + String.Format(fmt, s))
        WriteLine(vbTab & "ブロック (個)" & vbTab & ": " + String.Format(fmt, block))
        ' Credit
        If Show("続けたいですか？", "質問", YesNo, Question) = DialogResult.Yes Then
            RunApp()
        End If
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Create directory advanced.
    ''' </summary>
    ''' <param name="path">Folder path.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Delete file advanced.
    ''' </summary>
    ''' <param name="path">File path.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Convert to G.
    ''' </summary>
    ''' <param name="num">Number.</param>
    ''' <returns>Number converted.</returns>
    Private Function ConvertToG(num As Double)
        Return If(num < 30, num * 910, num)
    End Function

    ''' <summary>
    ''' Format number.
    ''' </summary>
    ''' <param name="c">Circuit.</param>
    ''' <param name="s">Spread.</param>
    ''' <param name="block">Block.</param>
    ''' <returns>Format.</returns>
    Private Function FmtNo(c As Double, s As Double, block As Double)
        Dim cSize = c.ToString().Length
        Dim sSize = s.ToString().Length
        Dim blockSize = block.ToString().Length
        Dim maxSize = Max(blockSize, Max(cSize, sSize))
        Return "{0," + maxSize.ToString() + ":####.#}"
    End Function
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr.Enabled Then
            tmr.Start()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' Intro.
    ''' </summary>
    Private Sub Intro()
        Clear()
        ForegroundColor = Blue
        WriteLine(My.Resources.gr_name)
        WriteLine(My.Resources.cc_text)
        ForegroundColor = Green
        WriteLine(vbCrLf & My.Resources.app_true_name & vbCrLf)
    End Sub

    ''' <summary>
    ''' Header double input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function HdrDInp(caption As String)
        Intro()
        ForegroundColor = Cyan
        Write(caption)
        ForegroundColor = White
        Return Val(ReadLine)
    End Function
#End Region
End Module
