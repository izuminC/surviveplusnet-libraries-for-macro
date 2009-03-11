Imports System.Runtime.CompilerServices

''' <summary>
''' ListDialogWindow に表示する項目を司る COMクラスです。
''' VBA や VBScript から、参照設定、あるいは CreateObject で "Net.Surviveplus.LibrariesForMacro.ListDialogItem" のインスタンスを作成することが出来ます。
''' </summary>
''' <remarks></remarks>
<ComClass(ListDialogItem.ClassId, ListDialogItem.InterfaceId, ListDialogItem.EventsId)> _
Public Class ListDialogItem

#Region "COM GUID"
    ' これらの GUID は、このクラスおよびその COM インターフェイスの COM ID を 
    ' 指定します。この値を変更すると、 
    ' 既存のクライアントはクラスにアクセスできなくなります。
    Public Const ClassId As String = "970472ce-c135-4ea4-91a0-6cc2d8f6a06e"
    Public Const InterfaceId As String = "831516aa-0ece-44a3-95e5-ce6d40546954"
    Public Const EventsId As String = "06ef4705-14e8-4345-9751-bc31d97ba299"
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

    ''' <summary>
    ''' Name プロパティのバッキングフィールドです。
    ''' </summary>
    ''' <remarks></remarks>        
    Private valueOfName As String

    ''' <summary>
    ''' 項目の名称を取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Name() As String
        Get
            Return Me.valueOfName
        End Get
        Set(ByVal value As String)
            Me.valueOfName = value
        End Set
    End Property


    ''' <summary>
    ''' Explanation プロパティのバッキングフィールドです。
    ''' </summary>
    ''' <remarks></remarks>        
    Private valueOfExplanation As String

    ''' <summary>
    ''' 項目の説明を取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Explanation() As String
        Get
            Return Me.valueOfExplanation
        End Get
        Set(ByVal value As String)
            Me.valueOfExplanation = value
        End Set
    End Property


    ''' <summary>
    ''' Key プロパティのバッキングフィールドです。
    ''' </summary>
    ''' <remarks></remarks>        
    Private valueOfKey As String

    ''' <summary>
    ''' 項目にアクセスするためのキーを表す文字列を取得または設定します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Key() As String
        Get
            Return Me.valueOfKey
        End Get
        Set(ByVal value As String)
            Me.valueOfKey = value
        End Set
    End Property

#End Region

End Class

''' <summary>
''' ListDialogItem に拡張メソッドを追加するためのモジュールです。
''' </summary>
''' <remarks>
''' ListDialogItem クラスはCOMクラスなので、COM 公開の Public メンバのみを定義するようにして、
''' COM 非公開の追加機能は、このモジュールに拡張メソッドとして追加します。
''' </remarks>
Public Module ListDialogItemExtensions

#Region " GetKey / SetKey メソッド "

    ''' <summary>
    ''' 項目にアクセスするためのキーを表す System.Window.Input.Key 列挙値を取得します。
    ''' </summary>
    ''' <param name="item"></param>
    ''' <returns>Key プロパティに該当する System.Window.Input.Key 列挙値を取得します。該当する値が無い場合は None を返します。</returns>
    ''' <remarks></remarks>
    <Extension()> Public Function GetKey(ByVal item As ListDialogItem) As Key
        Try
            Dim keyText = item.Key
            If IsNumeric(keyText) Then
                keyText = "D" & keyText
            End If
            Return [Enum].Parse(GetType(Key), keyText)
        Catch
            Return Input.Key.None
        End Try
    End Function

    ''' <summary>
    ''' 項目にアクセスするためのキーを、System.Window.Input.Key 列挙値で指定して設定します。
    ''' </summary>
    ''' <param name="item"></param>
    ''' <param name="newKey">新しいキーを指定します。</param>
    ''' <remarks></remarks>
    <Extension()> Public Sub SetKey(ByVal item As ListDialogItem, ByVal newKey As Key)
        item.Key = newKey.ToString()
    End Sub
#End Region

End Module


