''' <summary>
''' リストダイアログ ウィンドウを司るクラスです。
''' 使用するには COMクラス（ListDialogクラス）を CreateObject してください。
''' </summary>
''' <remarks>
''' </remarks>
Partial Public Class ListDialogWindow

#Region " プロパティ "

    ''' <summary>
    ''' Items プロパティのバッキングフィールドです。
    ''' </summary>
    ''' <remarks></remarks>        
    Private valueOfItems As New List(Of ListDialogItem)

    ''' <summary>
    ''' ダイアログに表示する項目のリストを取得します。
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Items() As List(Of ListDialogItem)
        Get
            Return Me.valueOfItems
        End Get
    End Property


    ''' <summary>
    ''' 選択された項目を取得します。項目が選択されていないときは Nothing (C# の場合は null 参照) を返します。
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedItem() As ListDialogItem
        Get
            If Me.List.SelectedIndex >= 0 Then
                Return Me.Items(Me.List.SelectedIndex)
            Else
                Return Nothing
            End If
        End Get
    End Property

#End Region

#Region " イベント処理 "

    ''' <summary>
    ''' ウィンドウが初期化されたときの処理を実行します。
    ''' Items プロパティを List（ListViewコントロール）にデータバインドします。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListDialogWindow_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

        Me.List.DataContext = Me.valueOfItems
    End Sub

    ''' <summary>
    ''' 要素の配置、描画、および操作の準備が完了したときに発生します。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListDialogWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        ' フォームを最前面に表示することを試みます。（特に Visual Studio マクロから呼び出した時は、自らアクティブにならないといけない。
        Me.Activate()
    End Sub


    ''' <summary>
    ''' OK ボタンが押されたときの処理を実行します。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OkButton.Click

        Me.ExecuteOk()
    End Sub

    ''' <summary>
    ''' OKボタンが押されたときの内部処理を実行します。
    ''' 項目が選択されているときはウィンドウを閉じます。選択されていないときは何もしません。
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExecuteOk()

        If Me.List.SelectedIndex >= 0 Then
            Me.DialogResult = True
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' キャンセルボタンが押されたときの処理を実行します。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CancelButton.Click

    End Sub

    ''' <summary>
    ''' キーが押された時の処理を実行します。
    ''' 項目に指定されたキーと一致するときは、その項目を選択して OK ボタンを押してウィンドウを閉じます。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Window1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles MyBase.KeyDown

        If e.Key <> Key.None Then

            Dim index As Integer = 0
            For Each item As ListDialogItem In Me.Items
                Dim toKey = item.GetKey()
                If toKey = e.Key Then
                    Me.List.SelectedIndex = index
                    Me.ExecuteOk()
                    Exit For
                End If
                index += 1
            Next item
        End If

    End Sub

    ''' <summary>
    ''' マウスがダブルクリックされた時の処理を実行します。
    ''' OKボタンを押してウィンドウを閉じます。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub List_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles List.MouseDoubleClick
        Me.ExecuteOk()
    End Sub

#End Region

End Class
