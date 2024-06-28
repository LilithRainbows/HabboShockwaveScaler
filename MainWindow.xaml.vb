Imports System.IO
Imports System.Text
Imports Microsoft.Win32

Class MainWindow
    Public CurrentLanguageInt As Integer = 0
    Public MaxClientVersion = 1000
    Public InstallIdentifier = 1

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If Globalization.CultureInfo.CurrentCulture.Name.ToLower.StartsWith("es") Then
            CurrentLanguageInt = 1
        End If
        If Globalization.CultureInfo.CurrentCulture.Name.ToLower.StartsWith("pt") Then
            CurrentLanguageInt = 2
        End If
        Directory.SetCurrentDirectory(GetExecutableDirectory)
        AboutButton.Content = AppTranslator.AboutTitle(CurrentLanguageInt)
        UpdateScalingButton()
    End Sub

    Function GetExecutableDirectory() As String
        Return Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
    End Function

    Private Sub ToggleScalingButton_Click(sender As Object, e As RoutedEventArgs) Handles ToggleScalingButton.Click
        If ToggleScalingButton.Content = AppTranslator.EnableScaling(CurrentLanguageInt) Then
            If MessageBox.Show(AppTranslator.BugWarningContent(CurrentLanguageInt), AppTranslator.BugWarningTitle(CurrentLanguageInt), MessageBoxButton.YesNo, MessageBoxImage.Warning) = MessageBoxResult.Yes Then
                EnableScaling()
            End If
        Else
            DisableScaling()
        End If
        UpdateScalingButton()
    End Sub

    Private Sub AboutButton_Click(sender As Object, e As RoutedEventArgs) Handles AboutButton.Click
        MsgBox(AppTranslator.AboutContent(CurrentLanguageInt), MsgBoxStyle.Information, AppTranslator.AboutTitle(CurrentLanguageInt))
    End Sub

    Public Sub UpdateScalingButton()
        If CheckScalingStatus() Then
            ToggleScalingButton.Content = AppTranslator.DisableScaling(CurrentLanguageInt)
        Else
            ToggleScalingButton.Content = AppTranslator.EnableScaling(CurrentLanguageInt)
        End If
    End Sub

    Public Function CheckScalingStatus() As Boolean
        Try
            Dim InstalledIdentifierPath = Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\identifier"
            If File.Exists(InstalledIdentifierPath) Then
                If File.ReadAllText(InstalledIdentifierPath) = InstallIdentifier Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox("Error while checking scaling status.", MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function EnableScaling() As Boolean
        Try
            'Copy IntegerScaler embedded files
            KillPossibleIntegerScalerExecutable()
            Directory.CreateDirectory(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath))
            File.WriteAllBytes(GetPossibleIntegerScalerExecutablePath, GetIntegerScalerExecutableEmbeddedResource)
            File.WriteAllText(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\IntegerScaler_README.txt", My.Resources.IntegerScaler_README, Encoding.UTF8)
            File.WriteAllText(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\settings", "")
            'Make habbo shockwave client auto-scaled (supported versions: 1 to 1000)
            Dim AutoScaleSettings As String = ""
            For PossibleClientVersion = 1 To MaxClientVersion
                AutoScaleSettings += Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Habbo Launcher\downloads\shockwave\" & PossibleClientVersion & "\Habbo.exe" & Environment.NewLine
            Next
            File.WriteAllText(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\auto.txt", AutoScaleSettings, Encoding.UTF8)
            'Make habbo shockwave client dpi aware (supported versions: 1 to 1000)
            Using key = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers")
                For PossibleClientVersion = 1 To MaxClientVersion
                    key.SetValue(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Habbo Launcher\downloads\shockwave\" & PossibleClientVersion & "\Habbo.exe", "~ HIGHDPIAWARE", RegistryValueKind.String)
                Next
            End Using
            'Make HabboShockwaveScaler start with windows
            Dim InstalledHabboShockwaveScalerPath = Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\" & Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName)
            IO.File.Copy(Process.GetCurrentProcess().MainModule.FileName, InstalledHabboShockwaveScalerPath, True)
            AddAppToStartup("HabboShockwaveScaler", InstalledHabboShockwaveScalerPath, "-LaunchIntegerScaler")
            'Generate install identifier
            File.WriteAllText(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath) & "\identifier", InstallIdentifier)
            'Restart IntegerScaler process
            KillPossibleIntegerScalerExecutable()
            StartPossibleIntegerScalerExecutable()
            Return True
        Catch ex As Exception
            MsgBox(AppTranslator.EnableScalingError(CurrentLanguageInt), MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function DisableScaling() As Boolean
        Try
            'Remove IntegerScaler local files
            If Directory.Exists(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath)) Then
                KillPossibleIntegerScalerExecutable()
                Directory.Delete(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath), True)
            End If
            'Remove all habbo shockwave client dpi aware references
            Using key = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers")
                For Each keyname In key.GetValueNames
                    If keyname.Contains("\Habbo Launcher\downloads\shockwave\") And keyname.EndsWith("\Habbo.exe") Then
                        key.DeleteValue(keyname)
                    End If
                Next
            End Using
            'Remove HabboShockwaveScaler windows startup references
            Using key = Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
                For Each keyname In key.GetValueNames
                    If key.GetValue(keyname).ToString.Contains(Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath)) Then
                        key.DeleteValue(keyname)
                    End If
                Next
            End Using
            Return True
        Catch ex As Exception
            MsgBox(AppTranslator.DisableScalingError(CurrentLanguageInt), MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Sub StartPossibleIntegerScalerExecutable()
        Using NewIntegerScalerProcess As New Process
            NewIntegerScalerProcess.StartInfo.FileName = GetPossibleIntegerScalerExecutablePath()
            NewIntegerScalerProcess.StartInfo.Arguments = "-fractional -clipcursor -nohotkeys"
            NewIntegerScalerProcess.StartInfo.WorkingDirectory = Path.GetDirectoryName(GetPossibleIntegerScalerExecutablePath)
            NewIntegerScalerProcess.Start()
        End Using
    End Sub

    Public Sub KillPossibleIntegerScalerExecutable()
        For Each PossibleIntegerScalerExecutable In Process.GetProcessesByName(Path.GetFileNameWithoutExtension(GetPossibleIntegerScalerExecutablePath))
            If PossibleIntegerScalerExecutable.MainModule.FileName = GetPossibleIntegerScalerExecutablePath() Then
                PossibleIntegerScalerExecutable.Kill()
                PossibleIntegerScalerExecutable.WaitForExit()
            End If
        Next
    End Sub

    Public Sub AddAppToStartup(appName As String, appPath As String, appArguments As String)
        Using key = Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
            If String.IsNullOrWhiteSpace(appArguments) = False AndAlso appArguments.StartsWith(" ") = False Then
                appArguments = " " & appArguments
            End If
            key.SetValue(appName, Chr(34) & appPath & Chr(34) & appArguments)
        End Using
    End Sub

    Public Function GetPossibleIntegerScalerExecutablePath() As String
        If Environment.Is64BitOperatingSystem Then
            Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\HabboShockwaveScaler\IntegerScaler_x64.exe"
        Else
            Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\HabboShockwaveScaler\IntegerScaler_x86.exe"
        End If
    End Function

    Public Function GetIntegerScalerExecutableEmbeddedResource() As Byte()
        If Environment.Is64BitOperatingSystem Then
            Return My.Resources.IntegerScaler_x64
        Else
            Return My.Resources.IntegerScaler_x86
        End If
    End Function
End Class

Public Class AppTranslator
    '0=English 1=Spanish 2=Portuguese
    Public Shared AboutTitle As String() = {
        "About",
        "Acerca",
        "Sobre"
    }
    Public Shared AboutContent As String() = {
        "Made possible thanks to Marat Tanalin's IntegerScaler technology.",
        "Hecho posible gracias a la tecnología IntegerScaler de Marat Tanalin.",
        "Ele é possível graças à tecnologia IntegerScaler de Marat Tanalin."
    }
    Public Shared EnableScaling As String() = {
        "Enable scaling",
        "Habilitar escalado",
        "Habilitar escalonamento"
    }
    Public Shared DisableScaling As String() = {
        "Disable scaling",
        "Deshabilitar escalado",
        "Desativar escalonamento"
    }
    Public Shared EnableScalingError As String() = {
        "Could not enable scaling.",
        "No se pudo habilitar el escalado.",
        "Não foi possível habilitar o escalonamento."
    }
    Public Shared DisableScalingError As String() = {
        "Could not disable scaling.",
        "No se pudo deshabilitar el escalado.",
        "Não foi possível desativar o escalonamento."
    }
    Public Shared BugWarningTitle As String() = {
        "Bug warning",
        "Advertencia de bug",
        "Advertência de bug"
    }
    Public Shared BugWarningContent As String() = {
        "On some older incompatible devices, scaling may cause you to have to abruptly close your windows session in order to close the client." & Environment.NewLine & "Just in case, it is recommended that you save any unsaved work before continuing." & Environment.NewLine & "In case this bug happens to you, close your windows session using ctrl+alt+supr and then disable scaling using this application." & Environment.NewLine & Environment.NewLine & "Do you want to continue?",
        "En algunos dispositivos antiguos incompatibles el escalado puede ocasionar que tengas que cerrar tu sesion de windows de forma brusca para poder cerrar el cliente." & Environment.NewLine & "Por las dudas se recomienda que guardes cualquier trabajo no guardado antes de continuar." & Environment.NewLine & "En caso que te ocurra ese bug cierra tu sesion de windows usando ctrl+alt+supr y luego deshabilita el escalado usando esta aplicacion." & Environment.NewLine & Environment.NewLine & "Deseas continuar?",
        "Em alguns dispositivos incompatíveis mais antigos, o escalonamento pode fazer com que você feche abruptamente a sessão do windows para fechar o cliente." & Environment.NewLine & "Caso esse bug aconteça com você, feche sua sessão do windows usando ctrl+alt+supr e desative o escalonamento usando este aplicativo." & Environment.NewLine & Environment.NewLine & "Você quer continuar?"
    }
End Class