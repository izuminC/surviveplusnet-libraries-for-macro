Partial Public Class InputDialogWindow


#Region " プロパティ "

    ''' <summary>
    ''' ユーザーが入力したテキストを取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Text() As String
        Get
            Return Me.inputBox.Text
        End Get
        Set(ByVal value As String)
            Me.inputBox.Text = value
        End Set
    End Property


    ''' <summary>
    ''' AllowEmpty プロパティのバッキングフィールドです。
    ''' </summary>
    ''' <remarks></remarks>        
    Private valueOfAllowEmpty As Boolean

    ''' <summary>
    ''' 入力値に空文字を指定できるかどうかを取得または設定します。
    ''' </summary>
    ''' <remarks>
    ''' True を設定したときは、入力値に何も入力しなくても OK ボタンを押してウィンドウを閉じることが出来ます。
    ''' False を設定したときは、入力値に何か入力されなければ OK ボタンを押してもウィンドウを閉じることが出来ません。
    ''' 規定値は False です。
    ''' </remarks>
    Public Property AllowEmpty() As Boolean
        Get
            Return Me.valueOfAllowEmpty
        End Get
        Set(ByVal value As Boolean)
            Me.valueOfAllowEmpty = value
        End Set
    End Property


#End Region

#Region " イベント処理 "

    ''' <summary>
    ''' ウィンドウが初期化されたときの処理を実行します。
    ''' 入力ボックスにフォーカスを移し、エラー表示をクリアします。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InputDialogWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.inputBox.SelectAll()
        Me.inputBox.Focus()

        Me.errorMessage.Content = String.Empty
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
        If Me.AllowEmpty OrElse _
            String.IsNullOrEmpty(Me.Text) = False Then

            Me.errorMessage.Content = String.Empty
            Me.DialogResult = True
            Me.Close()
        Else
            Me.errorMessage.Content = "* 値を入力してください。"
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

#End Region

End Class
