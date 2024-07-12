Public Class frmConf

    Private Sub frmConf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sh As Object
        sh = CreateObject("WScript.Shell")
        tbdirftp.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbdirftp")
        tftplibreria.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\FTPLIB")
        tbcarpftp.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbcarpftp")
        tbuserftp.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbuserftp")
        tbpassftp.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbpassftp")
        emailconf.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Emailftp")
        tbcarpetaconf.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta")
        emailin1.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter1")
        emailin2.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter2")
        emailin3.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter3")
        emailVIA.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiASG")
        carpprocesados.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpProcesados")
        CarpErr.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpERR")
        carpmanual.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpManual")
        numasigria.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NUMASG")
        thoramaxima.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\HORAMAX")
        thorainicio.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\HORAINIC")
        tsabadomaxima.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\SABAMAX")
        nomarchs.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NOMARCHS")
        carpformateados.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpFormateados")
        libdown.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\FTPLibDown")
        carpdown.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\FTPCarpDown")
        NomAplicadas.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NomArchAplic")
        NomCancelaciones.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NomArchCance")
        NomDiario.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NomArchRDiario")
        CAplicadas.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Aplicadas")
        CCancelaciones.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Cancelaciones")
        CarpSal.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Salidas")
        CDiario.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Diario")
        CarpLogs.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpLogs")
        TxtCiclo.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CicloTime")
        TxtASGIp.Text = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\ASGIp") 'Ip SFTP de ASG
        TxtASGPass.Text = sh.Regread("HKEY_CURRENT_USER\Software\BancocciASG\ASGPass") 'usuario SFTP de ASG
        TxtASGUser.Text = sh.Regread("HKEY_CURRENT_USER\Software\BancocciASG\ASGUser") 'password SFTP de ASG
        sh = Nothing

    End Sub

    Private Sub btnGuardarConf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardarConf.Click
        Dim sh As Object
        sh = CreateObject("WScript.Shell")
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbdirftp", tbdirftp.Text) 'Direccion IP FTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPLIB", tftplibreria.Text) 'Libreria FTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbcarpftp", tbcarpftp.Text) 'Carpeta FTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbuserftp", tbuserftp.Text) 'USER FTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbpassftp", tbpassftp.Text) 'PASS FTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta", tbcarpetaconf.Text) 'PATH Carpeta del DAEMON
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Emailftp", emailconf.Text) 'Primer Correo (FTP)
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter1", emailin1.Text) 'Segundo Correo
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter2", emailin2.Text) 'Tercer Correo
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter3", emailin3.Text) 'Cuarto Correo
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiASG", emailVIA.Text) 'Email Vi Americas
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpProcesados", carpprocesados.Text) 'Carpeta Archivos Procesados
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpManual", carpmanual.Text) 'Carpeta Archivos Manual
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpERR", CarpErr.Text) 'Carpeta Archivos Con Error
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NUMASG", numasigria.Text) 'Numero de Asignacion de Via
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\HORAMAX", thoramaxima.Text) 'Hora maxima de revisión de Lunes a Viernes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\HORAINIC", thorainicio.Text) 'Hora inicio de revisión de Lunes a Viernes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\SABAMAX", tsabadomaxima.Text) 'Hora maxima de revisión Sabados
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NOMARCHS", nomarchs.Text) 'Nombre de Archivos 2 Posiciones
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpFormateados", carpformateados.Text) 'Carpeta de Archivos Formateados
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Salida", CarpSal.Text) 'Carpeta de Archivos de Salida por SFTP
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPLibDown", libdown.Text) 'Libreria FTP para Bajar Reportes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPCarpDown", carpdown.Text) 'Carpeta FTP para Bajar Reportes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchAplic", NomAplicadas.Text) 'Carpeta de Archivos Formateados
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchCance", NomCancelaciones.Text) 'Libreria FTP para Bajar Reportes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchRDiario", NomDiario.Text) 'Carpeta FTP para Bajar Reportes
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpLogs", CarpLogs.Text) 'Carpeta LOGS
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGIp", TxtASGIp.Text) 'Ip SFTP de ASG
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGPass", TxtASGPass.Text) 'usuario SFTP de ASG
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGUser", TxtASGUser.Text) 'password SFTP de ASG
        sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CicloTime", TxtCiclo.Text) 'password SFTP de ASG

        frmPrincipalASG.CicloCarp = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta")
        sh = Nothing
        MsgBox("Debe reiniciar el programa para que algunos cambios tengan efecto")
    End Sub

    Private Sub btnCancelConf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelConf.Click
        Me.Dispose()
    End Sub

    Private Sub btnexamcarpconf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexamcarpconf.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                tbcarpetaconf.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                carpprocesados.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                carpmanual.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                CarpErr.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                carpformateados.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                CarpLogs.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim oFolderBrowser As New FolderBrowserDialog
        With oFolderBrowser
            .SelectedPath = ""
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                CarpSal.Text = .SelectedPath & "\"
                .Dispose()
            End If
        End With
    End Sub
End Class