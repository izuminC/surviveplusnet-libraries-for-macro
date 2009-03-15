''' <summary>
''' クリップボード を COM 相互運用で使用できるように公開するための、COMクラスです。
''' VBA や VBScript から、参照設定、あるいは CreateObject で "Net.Surviveplus.LibrariesForMacro.Clipboard" のインスタンスを作成することが出来ます。
''' </summary>
''' <remarks>
''' <para>
''' Visual Studio から COMクラスをビルドするには、レジストリに登録するために管理者権限が必要です。
''' Windows Vista 以降では、Visual Studio を管理者として実行してください。
''' </para>
''' </remarks>
<ComClass(Clipboard.ClassId, Clipboard.InterfaceId, Clipboard.EventsId)> _
Public Class Clipboard

#Region "COM GUID"
    ' これらの GUID は、このクラスおよびその COM インターフェイスの COM ID を 
    ' 指定します。この値を変更すると、 
    ' 既存のクライアントはクラスにアクセスできなくなります。
    Public Const ClassId As String = "0442583d-a84b-44b0-bdce-141e45b59458"
    Public Const InterfaceId As String = "7c23eedb-43b4-4dd3-8903-20aafd117c3f"
    Public Const EventsId As String = "a097b83e-d86c-4bb0-831b-799e165afa53"
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

    ''' <summary>
    ''' System.Windows.Clipboard.GetText のラッパーメソッドです。
    ''' クリップボードのテキストを取得して返します。
    ''' </summary>
    ''' <returns>クリップボードのテキストを返します。</returns>
    ''' <remarks></remarks>
    Public Function GetText() As String

        Return System.Windows.Clipboard.GetText()
    End Function

    ''' <summary>
    ''' System.Windows.Clipboard.SetText のラッパーメソッドです。
    ''' クリップボードにテキストを格納します。
    ''' </summary>
    ''' <param name="text">クリップボードに格納するテキストを指定します。</param>
    ''' <remarks></remarks>
    Public Sub SetText(ByVal text As String)

        System.Windows.Clipboard.SetText(text)
    End Sub

    ''' <summary>
    ''' System.Windows.Clipboard.GetFileDropList のラッパーメソッドです。
    ''' クリップボードから取得できるドロップされたファイルのリストを含む文字列を返します。
    ''' </summary>
    ''' <returns>クリップボードのファイルリストを、１行に１ファイルを列挙したテキストとして連結して返します。ファイル名はフルパスです。</returns>
    ''' <remarks></remarks>
    Public Function GetFileDropListText() As String
        Dim result As New System.Text.StringBuilder()

        For Each fileName In System.Windows.Clipboard.GetFileDropList()
            result.AppendLine(fileName)
        Next
        Return result.ToString()
    End Function

    ''' <summary>
    ''' System.Windows.Clipboard.SetFileDropList のラッパーメソッドです。
    ''' クリップボードにファイルリストをコピーします。コピーしたファイルはエクスプローラ等で貼り付ける事が出来ます。
    ''' </summary>
    ''' <param name="filesText">クリップボードにコピーするファイルリストを、1行に1ファイルを列挙したテキストとして連結して指定します。</param>
    ''' <remarks></remarks>
    Public Sub SetFileDropyListText(ByVal filesText As String)

        Dim files As New System.Collections.Specialized.StringCollection()
        For Each fileName In filesText.Split(vbCrLf)
            files.Add(Trim(fileName))
        Next
        System.Windows.Clipboard.SetFileDropList(files)
    End Sub

End Class


