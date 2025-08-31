Imports Microsoft.Web.WebView2.Core
Imports System.IO
Imports System.Reflection.Metadata
Public Class FamilySearchWeb

    Private ReadOnly userDataFolder As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "inoGen",
                     "WebView2Data")

    Private targetUrl As String = "https://www.familysearch.org/tree/"
    Public Sub New(startUrl As String)
        InitializeComponent()
        InitBrowser(startUrl)
    End Sub

    Private Sub InitBrowser(startUrl As String)
        targetUrl = startUrl
    End Sub

    Private Sub Go_Click(sender As Object, e As RoutedEventArgs)
        If Not String.IsNullOrWhiteSpace(AddressBar.Text) Then
            Try
                Dim url = AddressBar.Text
                If Not url.StartsWith("http") Then
                    url = "https://" & url
                End If
                WebView.Source = New Uri(url)
            Catch ex As Exception
                MessageBox.Show("Ungültige URL: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub NavigationStartingHandler(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
        AddressBar.Text = e.Uri
    End Sub

    Private Sub NavigationCompletedHandler(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If e.IsSuccess = False Then
            MessageBox.Show("Fehler beim Laden der Seite.")
        End If
    End Sub

    Private Async Sub FamilySearchWeb_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
        Await WebView.EnsureCoreWebView2Async(env)

        Dim hasSession As Boolean = Directory.Exists(userDataFolder) AndAlso
                                    Directory.GetFiles(userDataFolder, "*", SearchOption.AllDirectories).Length > 0

        If hasSession Then
            WebView.CoreWebView2.Navigate(targetUrl)
        Else
            WebView.CoreWebView2.Navigate("https://www.familysearch.org/auth/familysearch/login")

            AddHandler WebView.CoreWebView2.NavigationCompleted,
                Sub(s, args)
                    If WebView.Source.AbsoluteUri.Contains("familysearch.org") AndAlso
                       Not WebView.Source.AbsoluteUri.Contains("/auth/familysearch/login") Then
                        WebView.CoreWebView2.Navigate(targetUrl)
                    End If
                End Sub
        End If

        AddHandler WebView.CoreWebView2.NavigationStarting, AddressOf NavigationStartingHandler
        AddHandler WebView.CoreWebView2.NavigationCompleted, AddressOf NavigationCompletedHandler
    End Sub

    Public Sub NavigateTo(url As String)
        If WebView.CoreWebView2 IsNot Nothing Then
            WebView.CoreWebView2.Navigate(url)
        Else
            targetUrl = url
        End If
    End Sub
End Class
