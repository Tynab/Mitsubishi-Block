﻿Imports System.Console
Imports System.Text.Encoding

Public Module ModuleMain
    ''' <summary>
    ''' Main.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8
        If Not My.Settings.Chk_Key Then
            If InputBox("シリアルを入力", "ライセンスキー") = My.Resources.key_ser Then
                UpdVldLic()
                FstRunApp()
            Else
                ErrSty("ライセンスが間違っています。 終了するには、任意のキーを押してください...")
                ReadKey()
            End If
        Else
            FstRunApp()
        End If
    End Sub
End Module
