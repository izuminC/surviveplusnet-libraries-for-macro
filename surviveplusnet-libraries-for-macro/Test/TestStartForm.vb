''' <summary>
''' Net.Surviveplus.LibrariesForMacro COMクラスをテストするためのフォームです。
''' 登録された一連の動作を実行し、ユーザーインタフェイスなどが正しく動作することを目視で確認します。
''' </summary>
''' <remarks>
''' <para>
''' COM クラスの CreateObject ではなく、.net Framework のクラスとして参照設定されているため、
''' COM 相互運用のテストにはなりません。ここでは各クラスの機能のテストを行います。
''' また、COMクラスをビルドするために Visual Studio が管理者として実行されている場合、
''' デバッグではこのフォームも管理者として実行される点に注意してください。
''' </para>
''' </remarks>
Public Class TestStartForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim dialog As New Net.Surviveplus.LibrariesForMacro.ListDialog()
        dialog.Title = "テストタイトル"

        For i As Integer = 0 To 5
            Dim item As New Net.Surviveplus.LibrariesForMacro.ListDialogItem()
            item.Key = i.ToString()
            item.Name = "名前" & i.ToString()
            item.Explanation = "説明" & i.ToString()

            dialog.Add(item)
        Next i

        dialog.AddNewItem("名前A", "説明A", "A")

        Dim result = dialog.ShowDialog()
        Dim itemName As String = "（未選択）"
        If result Then
            itemName = dialog.SelectedItem.Name
        End If
        Call MsgBox(String.Format( _
               "戻り値：{0}" & vbCrLf & "選択項目：{1}", _
                 result, _
                 itemName _
               ))

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim dialog As New Net.Surviveplus.LibrariesForMacro.InputDialog()
        dialog.Title = "テストタイトル"
        dialog.Text = "規定値"
        'dialog.AllowEmpty = True
        If dialog.ShowDialog() Then
            MsgBox(dialog.Text)
        End If

    End Sub

    ''' <summary>
    ''' x64 環境では、dll をビルドした時のレジストリ登録では正しく動作しません。このボタンを押して登録すると、x64 環境でも CreateObject できるようになります。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try

            Dim registration As New Runtime.InteropServices.RegistrationServices()
            Dim ass = GetType(Net.Surviveplus.LibrariesForMacro.Clipboard).Assembly
            Dim result As Boolean = registration.RegisterAssembly(ass, Runtime.InteropServices.AssemblyRegistrationFlags.None)

            If result = False Then
                Throw New NotSupportedException("Regstration.RegsterAssembly に失敗しました。")
            End If

        Catch ex As Exception
            MsgBox("レジストリの登録に失敗しました。" + ex.ToString(), MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Try
            Dim registration As New Runtime.InteropServices.RegistrationServices()
            Dim ass = GetType(Net.Surviveplus.LibrariesForMacro.Clipboard).Assembly
            Dim result As Boolean = registration.UnregisterAssembly(ass)

            If result = False Then
                Throw New NotSupportedException("Regstration.UnregisterAssembly  に失敗しました。")
            End If

        Catch ex As Exception
            MsgBox("レジストリの登録に失敗しました。" + ex.ToString(), MsgBoxStyle.Exclamation)
        End Try
    End Sub
End Class
