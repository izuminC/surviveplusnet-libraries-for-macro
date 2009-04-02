<ComClass(InputDialog.ClassId, InputDialog.InterfaceId, InputDialog.EventsId)> _
Public Class InputDialog

#Region "COM GUID"
    ' これらの GUID は、このクラスおよびその COM インターフェイスの COM ID を 
    ' 指定します。この値を変更すると、 
    ' 既存のクライアントはクラスにアクセスできなくなります。
    Public Const ClassId As String = "4632ea43-ed45-40f7-abe1-1d536dbb6382"
    Public Const InterfaceId As String = "72604068-5c5e-48ac-b2c4-740ff26d4bd0"
    Public Const EventsId As String = "d3141103-d4bc-4510-8a1b-cfe37734fb7a"
#End Region

#Region " コンストラクタ "

    ' 作成可能な COM クラスにはパラメータなしの Public Sub New() を指定しなければ 
    ' なりません。これを行わないと、クラスは COM レジストリに登録されず、 
    ' CreateObject 経由で 
    ' 作成できません。
    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region " Public プロパティ "

    Dim dialog As New InputDialogWindow

    ''' <summary>
    ''' ウィンドウのタイトルを取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Title() As String
        Get
            Return Me.dialog.Title
        End Get
        Set(ByVal value As String)
            Me.dialog.Title = value
        End Set
    End Property


    ''' <summary>
    ''' ユーザーが入力したテキストを取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Text() As String
        Get
            Return Me.dialog.Text
        End Get
        Set(ByVal value As String)
            Me.dialog.Text = value
        End Set
    End Property


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
            Return Me.dialog.AllowEmpty
        End Get
        Set(ByVal value As Boolean)
            Me.dialog.AllowEmpty = value
        End Set
    End Property

#End Region

#Region " Public メソッド "

    ''' <summary>
    ''' ダイアログを表示し、ユーザーの操作が終わるまで処理を待ちます。
    ''' </summary>
    ''' <returns>
    ''' ユーザーがダイアログの OKボタン を押した時は True を返します。それ以外は False を返します。
    ''' </returns>
    ''' <remarks></remarks>
    Public Function ShowDialog() As Boolean
        Return dialog.ShowDialog()
    End Function
#End Region


End Class


