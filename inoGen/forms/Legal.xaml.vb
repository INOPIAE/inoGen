Imports System.Reflection
Imports System.IO
Public Class Legal
    Public Sub New()
        InitializeComponent()

        Dim packages = ReadNuGetPackages("inoGen.vbproj")
        lstPackages.ItemsSource = packages

        ' --- Lizenztext laden ---
        Dim licensePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "License")
        If File.Exists(licensePath) Then
            txtLicense.Text = File.ReadAllText(licensePath)
        Else
            txtLicense.Text = "Keine Lizenzdatei gefunden."
        End If
    End Sub

    Public Function ReadNuGetPackages(projectFile As String) As List(Of String)
        Dim result As New List(Of String)
        Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
        Dim projectDir = Path.GetFullPath(Path.Combine(baseDir, "..\..\..\"))
        projectFile = Path.Combine(projectDir, projectFile)

        If Not File.Exists(projectFile) Then
            result.Add("Projektdatei nicht gefunden: " & projectFile)
            Return result
        End If

        Dim doc = XDocument.Load(projectFile)

        Dim packages = From p In doc.Descendants("PackageReference")
                       Let inc = p.Attribute("Include")?.Value
                       Let ver = p.Attribute("Version")?.Value
                       Where inc IsNot Nothing And ver IsNot Nothing
                       Select New With {.Name = inc, .Version = ver}

        For Each pkg In packages
            Dim pkgPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".nuget", "packages", pkg.Name.ToLower(), pkg.Version
        )

            Dim nuspec = Directory.GetFiles(pkgPath, "*.nuspec", SearchOption.AllDirectories).FirstOrDefault()
            Dim license = "(Lizenz nicht gefunden)"

            If nuspec IsNot Nothing AndAlso File.Exists(nuspec) Then
                Dim nuspecDoc = XDocument.Load(nuspec)
                Dim licNode = nuspecDoc.Descendants().FirstOrDefault(Function(x) x.Name.LocalName = "license")
                Dim licUrl = nuspecDoc.Descendants().FirstOrDefault(Function(x) x.Name.LocalName = "licenseUrl")

                If licNode IsNot Nothing Then
                    license = licNode.Value
                ElseIf licUrl IsNot Nothing Then
                    license = licUrl.Value
                End If
            End If

            result.Add($"{pkg.Name} - {pkg.Version} - Lizenz: {license}")
        Next

        Return result
    End Function
End Class
