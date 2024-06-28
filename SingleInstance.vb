Imports System.IO
Module SingleInstance 'App framework must be disabled from project settings and then select SingleInstance as Startup Object 
    Sub Main(Args As String())
        Dim noPreviousInstance As Boolean
        Using m As New Threading.Mutex(True, "HabboShockwaveScaler", noPreviousInstance)
            If Not noPreviousInstance Then
                MessageBox.Show("Program is already running!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            Else
                Dim NewMainWindow As New MainWindow(Args)
                Dim NewApp As New Application()
                NewApp.Run(NewMainWindow)
            End If
        End Using
    End Sub
End Module
