Option Explicit On
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Permissions


Public Class Form1

    Dim WithEvents TabControl1 As New TabControl
    Dim WithEvents WebBrowser1 As ExWebBrowser
    Dim newTab As New TabPage
    Dim WithEvents TabPage1 As TabPage
    Dim setuptext As String = ""
    Dim logoutRef As Boolean = False

    'TabControl1用のコンテキストメニュー部分3列
    Dim WithEvents menuTab As New ContextMenu()
    Dim WithEvents menuItemAdd As New MenuItem()
    Dim WithEvents menuItemClose As New MenuItem()

    Private refreshButton1 As System.Windows.Forms.Button
    Private nextButton1 As System.Windows.Forms.Button
    Private WithEvents moveButton1 As System.Windows.Forms.Button
    Private urlCombo1 As System.Windows.Forms.ComboBox


    'newBrowser
    Dim WithEvents ExWebBrowser As New WebBrowser


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Memory_r.Release()

        'フォームの大きさ
        Me.Size = New Size(1200, 600)

        'URL表示のコンボボックス
        Me.urlCombo1 = New ComboBox
        Me.urlCombo1.Name = "urlCombo1"
        'サイズと位置を設定する
        Me.urlCombo1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom Or AnchorStyles.Right
        Me.urlCombo1.Location = New Point(180, 0)
        urlCombo1.Size = New Size(922, 20)
        'フォームに追加する
        Me.Controls.Add(Me.urlCombo1)

        'Buttonクラスのインスタンスを作成する
        Me.moveButton1 = New System.Windows.Forms.Button()
        'Buttonコントロールのプロパティを設定する
        Me.moveButton1.Name = "moveButton1"
        Me.moveButton1.Text = "移動ボタン"
        'サイズと位置を設定する
        Me.moveButton1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.moveButton1.Location = New Point(0, 25)
        'フォームに追加する
        Me.Controls.Add(Me.moveButton1)

        'Buttonクラスのインスタンスを作成する
        Me.refreshButton1 = New System.Windows.Forms.Button()
        'Buttonコントロールのプロパティを設定する
        Me.refreshButton1.Name = "refreshButton1"
        Me.refreshButton1.Text = "戻る"
        'サイズと位置を設定する
        Me.refreshButton1.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        'フォームに追加する
        Me.Controls.Add(Me.refreshButton1)

        'Buttonクラスのインスタンスを作成する
        Me.nextButton1 = New System.Windows.Forms.Button()
        'Buttonコントロールのプロパティを設定する
        Me.nextButton1.Name = "nextButton1"
        Me.nextButton1.Text = "進む"
        'サイズと位置を設定する
        Me.nextButton1.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Me.nextButton1.Location = New Point(90, 0)
        'フォームに追加する
        Me.Controls.Add(Me.nextButton1)

        'WebBrowser1
        Me.WebBrowser1 = New ExWebBrowser
        Me.WebBrowser1.Dock = DockStyle.Fill
        AddHandler WebBrowser1.NewWindow2, AddressOf WebBrowser_NewWindow2


        'TabPage1
        Me.TabPage1 = New TabPage("tab1")
        TabPage1.BackColor = Color.White
        TabPage1.Controls.Add(New Button())

        'TabControl
        Me.TabControl1.Dock = DockStyle.None
        Me.TabControl1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TabControl1.Location = New Point(20, 50)
        TabControl1.Size = New Size(1150, 500)
        Me.TabControl1.TabPages.Add(TabPage1)


        'Form1
        Me.Text = "amb-EXL"
        Me.Controls.Add(Me.TabControl1)



        '[挿入]と[削除]項目の設定
        menuItemAdd.Text = "タブの追加(未実装)"
        menuItemClose.Text = "タブを閉じる"

        'TabControl1用のコンテキストメニューに[タブの追加]と[タブを閉じる]を追加
        menuTab.MenuItems.Add(menuItemAdd)
        menuTab.MenuItems.Add(menuItemClose)

        '作成したコンテキストメニューをTabControl1に設定
        TabControl1.ContextMenu = menuTab

    End Sub

    Private Sub menuItemClose_Click(sender As Object, e As System.EventArgs) Handles menuItemClose.Click
        TabControl1.SelectedTab = TabPage1
        Console.WriteLine(TabControl1.TabIndex.ToString())
        TabControl1.TabPages.Remove(value:=newTab)
    End Sub

    ' メッセージボックスで [いいえ] を選択した場合は、フォームが閉じられるのをキャンセル
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If MessageBox.Show("ログアウトはしましたか？", "確認", MessageBoxButtons.YesNo) = DialogResult.No Then
            e.Cancel = True

        End If
    End Sub


    Private Sub moveButton1_Click(sender As Object, e As EventArgs) Handles moveButton1.Click

        'URL表示
        'Me.WebBrowser1.Navigate("http://www.ameba.jp/")
        WebBrowser1.Url = New Uri("http://www.ameba.jp/")
        Try
            '・新しいタブを動的に作る方法

            newTab.Controls.Add(WebBrowser1)
            TabControl1.TabPages.Add(newTab)
            'ボタンをおした時、タブ２へ自動切り替え
            TabControl1.SelectedTab = newTab
        Catch ex As Exception

        End Try

        'ieキャッシュ他削除
        Try
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start("RunDll32", "InetCpl.cpl,ClearMyTracksByProcess 4351")
            p.WaitForExit()
        Catch ex As Exception

        End Try

        'ログインリフレッシュ
        logoutRef = True
    End Sub


    Private Sub WebBrowser_NewWindow2(ByVal sender As Object, ByVal e As WebBrowserNewWindow2EventArgs)
        'WebBrowser1
        WebBrowser1 = New ExWebBrowser
        WebBrowser1.Dock = DockStyle.Fill
        AddHandler WebBrowser1.NewWindow2, AddressOf WebBrowser_NewWindow2
        'TabPage1
        TabPage1 = New TabPage(WebBrowser1.DocumentTitle)

        'TabPage1 = New TabPage("名無しTAB")

        TabPage1.Controls.Add(WebBrowser1)
        'TabControl
        TabControl1.Controls.Add(TabPage1)
        TabControl1.SelectedTab = TabPage1
        '新しいウィンドウが開くのを抑制
        e.ppDisp = Me.WebBrowser1.Application
        WebBrowser1.RegisterAsBrowser = True
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted


        'URL（ドキュメントコンプリート）
        Try
            urlCombo1.Text = WebBrowser1.Url.ToString()
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try

        Try
            Me.Text = e.Url.ToString() + "[DocumentCompleted]"

            If Me.Text.Contains("about:blank") Then
                Me.Text = WebBrowser1.Url.ToString + "[DocumentCompleted]"

            End If
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try


        'newtabにタイトル挿入
        newTab.Text = WebBrowser1.DocumentTitle

        'IDパスをテキストから読み込み
        Try
            Dim all As HtmlElementCollection = WebBrowser1.Document.All
            Dim forms As HtmlElementCollection = all.GetElementsByName("amebaId")
            Dim forms2 As HtmlElementCollection = all.GetElementsByName("password")
            Try
                ' Open the file using a stream reader.
                Using sr As New StreamReader("setup.txt")

                    'テキストからIDを取得
                    Dim amdID As String
                    'テキストからPASSを取得
                    Dim amdPASS As String

                    'コメント部分をスルーさせるコード
                    sr.ReadLine()

                    'テキスト→ID取得
                    amdID = sr.ReadLine()
                    'ID部分だけ取り除く
                    Dim _amdID As String = amdID.Replace("ID=", "")

                    'テキストからPASSを取得
                    amdPASS = sr.ReadLine()
                    Dim _amdPASS As String = amdPASS.Replace("PASS=", "")

                    '下記URLの場合ブーリアンを偽にする
                    If WebBrowser1.Url.ToString = "http://www.ameba.jp/" AndAlso logoutRef = True Then
                        logoutRef = False
                        'IDとパスの取得と自動ログイン
                        forms(0).InnerText = _amdID
                        Console.WriteLine(_amdID)
                        forms2(0).InnerText = _amdPASS
                        Console.WriteLine(_amdPASS)


                        'ログインする
                        WebBrowser1.Document.Forms(0).InvokeMember("submit")

                    End If

                    sr.Close()
                End Using

            Catch ex As Exception
                'Console.WriteLine("The file could not be read:")
                'Console.WriteLine(ex.Message)
            End Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub WebBrowser1_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles WebBrowser1.Navigating
        Try
            Me.Text = e.Url.ToString() + "[Navigating]"

            If Me.Text.Contains("about:blank") Then
                Me.Text = WebBrowser1.Url.ToString + "[Navigating]"

            End If
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try


    End Sub


End Class

Class ExWebBrowser
    Inherits WebBrowser

    'NewWindow2イベントの拡張
    Private cookie As AxHost.ConnectionPointCookie
    Private helper As WebBrowser2EventHelper

    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
    <System.Runtime.InteropServices.DispIdAttribute(200)>
    Public ReadOnly Property Application() As Object
        Get
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("Application", AxHost.ActiveXInvokeKind.PropertyGet)
            End If
            Return Me.ActiveXInstance.Application
        End Get
    End Property

    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
    <System.Runtime.InteropServices.DispIdAttribute(552)>
    Public Property RegisterAsBrowser() As Boolean
        Get
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("RegisterAsBrowser", AxHost.ActiveXInvokeKind.PropertyGet)
            End If
            Return Me.ActiveXInstance.RegisterAsBrowser
        End Get
        Set(ByVal value As Boolean)
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("RegisterAsBrowser", AxHost.ActiveXInvokeKind.PropertySet)
            End If
            Me.ActiveXInstance.RegisterAsBrowser = value
        End Set
    End Property

    <PermissionSetAttribute(SecurityAction.LinkDemand, Name:="FullTrust")>
    Protected Overrides Sub CreateSink()
        MyBase.CreateSink()
        helper = New WebBrowser2EventHelper(Me)
        cookie = New AxHost.ConnectionPointCookie(Me.ActiveXInstance, helper, GetType(DWebBrowserEvents2))
    End Sub

    <PermissionSetAttribute(SecurityAction.LinkDemand, Name:="FullTrust")>
    Protected Overrides Sub DetachSink()
        If cookie IsNot Nothing Then
            cookie.Disconnect()
            cookie = Nothing
        End If
        MyBase.DetachSink()
    End Sub

    Public Event NewWindow2 As WebBrowserNewWindow2EventHandler

    Protected Overridable Sub OnNewWindow2(ByVal e As WebBrowserNewWindow2EventArgs)
        RaiseEvent NewWindow2(Me, e)
    End Sub

    Private Class WebBrowser2EventHelper
        Inherits StandardOleMarshalObject
        Implements DWebBrowserEvents2

        Private parent As ExWebBrowser

        Public Sub New(ByVal parent As ExWebBrowser)
            Me.parent = parent
        End Sub

        Public Sub NewWindow2(ByRef ppDisp As Object, ByRef cancel As Boolean) Implements DWebBrowserEvents2.NewWindow2
            Dim e As New WebBrowserNewWindow2EventArgs(ppDisp)
            Me.parent.OnNewWindow2(e)
            ppDisp = e.ppDisp
            cancel = e.Cancel
        End Sub

    End Class

End Class

Public Delegate Sub WebBrowserNewWindow2EventHandler(ByVal sender As Object, ByVal e As WebBrowserNewWindow2EventArgs)

Public Class WebBrowserNewWindow2EventArgs
    Inherits System.ComponentModel.CancelEventArgs

    Private ppDispValue As Object

    Public Sub New(ByVal ppDisp As Object)
        Me.ppDispValue = ppDisp
    End Sub

    Public Property ppDisp() As Object
        Get
            Return ppDispValue
        End Get
        Set(ByVal value As Object)
            ppDispValue = value
        End Set
    End Property

End Class

<ComImport(), Guid("34A715A0-6587-11D0-924A-0020AFC7AC4D"),
InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
TypeLibType(TypeLibTypeFlags.FHidden)>
Public Interface DWebBrowserEvents2

    <DispId(DISPID.NEWWINDOW2)> Sub NewWindow2(
        <InAttribute(), OutAttribute(), MarshalAs(UnmanagedType.IDispatch)> ByRef ppDisp As Object,
        <InAttribute(), OutAttribute()> ByRef cancel As Boolean)

End Interface

Public Enum DISPID
    NEWWINDOW2 = 251
End Enum
