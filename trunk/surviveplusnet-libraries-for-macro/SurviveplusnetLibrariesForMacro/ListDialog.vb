''' <summary>
''' WPF ウィンドウ ListDialogWindow を COM 相互運用で使用できるように公開するための、COMクラスです。
''' VBA や VBScript から、参照設定、あるいは CreateObject で "Net.Surviveplus.LibrariesForMacro.ListDialog" のインスタンスを作成することが出来ます。
''' </summary>
''' <remarks>
''' <para>
''' Visual Studio から COMクラスをビルドするには、レジストリに登録するために管理者権限が必要です。
''' Windows Vista 以降では、Visual Studio を管理者として実行してください。
''' </para>
''' </remarks>
<ComClass(ListDialog.ClassId, ListDialog.InterfaceId, ListDialog.EventsId)> _
Public Class ListDialog

#Region "COM GUID"
    ' これらの GUID は、このクラスおよびその COM インターフェイスの COM ID を 
    ' 指定します。この値を変更すると、 
    ' 既存のクライアントはクラスにアクセスできなくなります。
    Public Const ClassId As String = "b50352e2-d0f6-47cc-91fe-e133d0a8a0ef"
    Public Const InterfaceId As String = "8409ccae-4f35-4924-9ea4-a19a30e7bf5f"
    Public Const EventsId As String = "b569981a-d5f8-4794-97c2-fb2bb2041a9a"
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

    Dim dialog As New ListDialogWindow

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
    ''' 選択された項目を取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedItem() As ListDialogItem
        Get
            Return Me.dialog.SelectedItem
        End Get
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

    ''' <summary>
    ''' 項目を新たに作成して追加します。
    ''' </summary>
    ''' <param name="name">項目の名称を指定します。</param>
    ''' <param name="explanation">項目の説明を指定します。省略可能です。</param>
    ''' <param name="key">項目にアクセスするためのキーを指定します。0～9 あるいは A～Z の文字で指定します。省略可能です。</param>
    ''' <returns>追加した項目の参照を返します。</returns>
    ''' <remarks></remarks>
    Public Function AddNewItem(ByVal name As String, Optional ByVal explanation As String = "", Optional ByVal key As String = "None") As ListDialogItem

        Dim item As New ListDialogItem
        With item
            .Name = name
            .Explanation = explanation
            .Key = key
        End With

        Me.dialog.Items.Add(item)
        Return item

    End Function

    ''' <summary>
    ''' 項目を追加します。
    ''' </summary>
    ''' <param name="item">追加する項目を指定します。</param>
    ''' <returns>追加した項目の参照を返します。</returns>
    ''' <remarks></remarks>
    Public Function Add(ByVal item As ListDialogItem) As ListDialogItem
        Me.dialog.Items.Add(item)
        Return item
    End Function

#End Region

End Class


