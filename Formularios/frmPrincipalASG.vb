Option Explicit On
Imports System.Array
Imports System.IO
Imports System.Data.OleDb
Imports System.Reflection
Imports System.Net
Imports System.Threading
Imports System.Net.Mail

Imports Renci.SshNet
Imports Renci.SshNet.Sftp

Public Class frmPrincipalASG
    Inherits System.Windows.Forms.Form
    Public pathnuevoarch As String
    Dim rarchivo As String
    Dim nomarchivo As String
    Dim tipoarchivo As String
    Dim email As String
    Public direccion, stringcorreo As String
    Dim carpeta, usuario, password, dirftp As String
    Public FlagEncabezado, FlagDetalle As Boolean
    Public Cola_Nombre_Arch(9999999) As String
    Public Cola_Direccion_Arch(9999999) As String
    Public Activado As Boolean = True
    Public horainiciocola, horafinalcola, CarpetaBackup, xdia As String
    Public ASGIp, ASGUser, ASGPass As String

    'Public log As String
    Public vDireccioIpAs400, vUsuarioAs400, vContraseñaAs400 As String
    Public StringErrCall As String
    Public CarpLogs, nomarchs As String
    Public pgmaplicacion, pgmcancelacion, CicloCarp, Manual, Salidas, Entradas, Procesados As String

    'Estructura básica para control de procesos ejecutados.
    Public Structure InfoProcesoTarea
        Dim Msg As String
        Dim getExecStatus As Boolean
    End Structure

    '***************************************
    'Desarrollado por Luis Manuel Rodriguez
    '***************************************

    Private Sub Cargar_Csv(ByVal lv As ListView, ByVal sPathCsv As String, ByVal aColumnHeader As String(), ByVal sDelimitador As String)
        Try
            ' verificar que la ruta sea correcta
            If File.Exists(sPathCsv) = False Then
                Logtxtbox.AppendText("No se encontró el archivo: " & sPathCsv)
                Exit Sub
            End If

        Catch ex As Exception
            Logtxtbox.AppendText("Error carga CSV : " & ex.Message.ToString & Chr(13))
        End Try
    End Sub

    ''' <summary>
    ''' Carga del fromulario principal y los valores de los campos originales
    ''' </summary>
    Private Sub Form1_Load(
        ByVal sender As System.Object,
        ByVal e As System.EventArgs) Handles MyBase.Load
        Logtxtbox.ForeColor = Color.Gray
        Dim sh As Object

        With NotifyIcon1
            .Icon = Me.Icon
        End With
        'Defino la Variable de chequeo si el programa se ha ejecutado o no por primera vez
        Dim chequeoprimeravez As String
        'limpio cualquier valor de la variable por locacion de memoria.
        chequeoprimeravez = ""
        'definimos un objeto vacio

        'EnvioMail(1)

        'Le damos Forma al Objeto jugando con API
        sh = CreateObject("WScript.Shell")
        'Intentamos Leer un Registro (el de Chequeo)

        Try
            chequeoprimeravez = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Chequeo")
            'Limpiamos el Objeto
            sh = Nothing
            'Agarramos el Handler de Excepciones y lo verificamos

        Catch ex As Exception
            'Si llegamos aqui es porque no existia en el registro la clave "Chequeo"
            'Por lo tanto verificamos nuevamente para luego proceder
            If chequeoprimeravez = "" Then
                'damos forma al objeto nuevamente(pues lo habiamos limpiado)
                sh = CreateObject("WScript.Shell")

                'Escribimos los Registros Iniciales. (Estos solo se escriben como standar, la primera vez que se ejecuta el Programa)
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbdirftp", "10.4.101.51") ' "10.4.101.51") 'Desarrollo 10.2.101.50
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPLIB", "QDLS")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbcarpftp", "BOCAP")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbuserftp", "UAPALIC") '"UAASG") '"APLSRVVAR") 'APLSRVVAR
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\tbpassftp", "U4P4L1C013") '"SECRETARIA") 'APLSRVVAR
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta", "C:\D.D. Bancocci Soft\Backups_ASG\CicloTemp\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Entradas", "C:\D.D. Bancocci Soft\Backups_ASG\Entradas\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Chequeo", "si")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Emailftp", "email@email.com")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter1", "email@email.com")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter2", "email@email.com")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter3", "email@email.com")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiASG", "email@email.com")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpProcesados", "C:\D.D. Bancocci Soft\Backups_ASG\Procesados\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Carpmanual", "C:\D.D. Bancocci Soft\Backups_ASG\Reporte Manual\") 'useless
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpERR", "C:\D.D. Bancocci Soft\Backups_ASG\ERR\") 'useless
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NUMASG", "  11")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\HORAMAX", "1500")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\HORAINIC", "009")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\SABAMAX", "1130")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NOMARCHS", "RT")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpFormateados", "C:\D.D. Bancocci Soft\Backups_ASG\Formateados\") 'useless
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\carpbackup", "C:\D.D. Bancocci Soft\Backups_ASG\Backup\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Salidas", "C:\D.D. Bancocci Soft\Backups_ASG\Salidas\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPLibDown", "QDLS")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\FTPCarpDown", "BOCAP")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchAplic", "DPTR11.TRF")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchCance", "DPTC11.TRF")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\NomArchRDiario", "DPTP11.TRF")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Aplicadas", "DPTCA011")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Cancelaciones", "DPTCN011")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\Diario", "DPTPG011")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CarpLogs", "C:\D.D. Bancocci Soft\Backups_ASG\Logs\")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGIp", "asg.com.do")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGUser", "telemed\userbanocci")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\ASGPass", "ASGbo528")
                sh.RegWrite("HKEY_CURRENT_USER\Software\BancocciASG\CicloTime", "18000000")
                'Limpiamos nuevamente
                sh = Nothing
            End If
        End Try

        'Damos forma nuevamente
        sh = CreateObject("WScript.Shell")
        ' Creamos el Sistem Watcher e implantamos las propiedades.
        CicloCarp = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta")
        CarpetaBackup = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\carpbackup")

        'Le damos Forma al Objeto jugando con API
        sh = CreateObject("WScript.Shell")
        ftplib = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\FTPLIB")
        carpeta = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbcarpftp")
        usuario = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbuserftp")
        password = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbpassftp")
        dirftp = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\tbdirftp")
        email = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Emailftp")
        NumRia = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NUMASG")
        Horamaxima = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\HORAMAX")
        HInicio = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\HORAINIC")
        SMaxima = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\SABAMAX")
        CarpLogs = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpLogs")
        nomarchs = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\NOMARCHS")
        pgmaplicacion = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\Aplicadas") 'useless
        pgmcancelacion = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\Cancelaciones") 'useless
        Entradas = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\Entradas")
        Salidas = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\Salidas")
        Manual = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\Carpmanual") 'useless
        ASGIp = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\ASGIp")
        ASGUser = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\ASGUser")
        ASGPass = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\ASGPass")
        Procesados = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpProcesados") 'Aqui estan los archivos del ciclo continuo
        Ciclo.Interval = sh.regRead("HKEY_CURRENT_USER\Software\BancocciASG\CicloTime") ' Damos inicio al ciclo
        sh = Nothing

        'Horarios normales de aplicación
        horainiciocola = HInicio
        horafinalcola = Horamaxima

        Logtxtbox.AppendText("Servicio Inicializado, esperando archivos" & Chr(13))
        Activado = True

        ' Creo las carpetas nesesarias para el trabajo
        Dim dirBackup As New System.IO.DirectoryInfo(CarpetaBackup)
        Dim dirCiclo As New System.IO.DirectoryInfo(CicloCarp)
        Dim dirEntradas As New System.IO.DirectoryInfo(Entradas)
        Dim dirSalidas As New System.IO.DirectoryInfo(Salidas)
        Dim dirCarpLogs As New System.IO.DirectoryInfo(CarpLogs)

        If Not dirBackup.Exists Then Directory.CreateDirectory(CarpetaBackup)
        If Not dirCiclo.Exists Then Directory.CreateDirectory(CicloCarp)
        If Not dirEntradas.Exists Then Directory.CreateDirectory(Entradas)
        If Not dirCarpLogs.Exists Then Directory.CreateDirectory(CarpLogs)


    End Sub

    ''' <summary>
    ''' Funcion de envio de correos despues de cada cierre
    ''' </summary>
    Public Sub EnvioMail(ByRef codigomensaje As Integer)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim emailftp, emailnoti1, emailnoti2, emailnoti3, emailnotiria As String

        Dim sh As Object
        sh = CreateObject("WScript.Shell")
        emailftp = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\Emailftp")
        emailnoti1 = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter1")
        emailnoti2 = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter2")
        emailnoti3 = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiInter3")
        emailnotiria = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\EmailNotiASG")
        sh = Nothing

        'CONFIGURACIÓN DEL STMP
        Dim _SMTP As New SmtpClient
        _SMTP.Credentials = New System.Net.NetworkCredential("cobranzas@tarjetasoccidente.hn", "eventos")
        _SMTP.Host = "mail.tarjetasoccidente.hn"
        _SMTP.Port = 25

        If emailftp <> "email@email.com" Then
            Select Case codigomensaje
                Case 1
                    If Len(emailftp) <> 0 And (emailftp) <> "email@email.com" Then
                        _Message.[To].Add(emailftp)
                    End If
                    If Len(emailnoti1) <> 0 And (emailnoti1) <> "email@email.com" Then
                        _Message.[To].Add(emailnoti1)
                    End If
                    If Len(emailnoti2) <> 0 And (emailnoti2) <> "email@email.com" Then
                        _Message.[To].Add(emailnoti2)
                    End If
                    If Len(emailnoti3) And (emailnoti3) <> "email@email.com" Then
                        _Message.[To].Add(emailnoti3)
                    End If

                    '_Message.[To].Add(emailnotiria)
                    _Message.From = New MailAddress("cobranzas@tarjetasoccidente.hn", "Sistema de Cobranzas - Bancocci", System.Text.Encoding.UTF8)
                    _Message.Subject = "Notificación de envio de archivo ASG - Banco de Occidente"
                    _Message.SubjectEncoding = System.Text.Encoding.UTF8
                    _Message.Priority = System.Net.Mail.MailPriority.High

                    ' ADICION DE DATOS ADJUNTOS

                    'If File.Exists(CarpProcesados & (Date.Now.ToString("dd-MM-yyyy")) & "\" & nomdownload) Then
                    '    Dim _File As String = CarpProcesados & (Date.Now.ToString("dd-MM-yyyy")) & "\" & nomdownload
                    '    Dim _Attachment As New Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet)
                    '    _Message.Attachments.Add(_Attachment)
                    'End If

                    _Message.IsBodyHtml = True
                    _Message.Body =
                        "<img style='width: 500px; height: 100px;' alt='LOGO BANCO'" & Chr(13) &
                        "title='LOGO'" & Chr(13) &
                        "src='http://www.bancocci.hn/images/header.png'><span" & Chr(13) &
                        "style='font-weight: bold;'><br>" & Chr(13) &
                        "<br>" & Chr(13) &
                        "Los Archivos del dia:&nbsp; </span><span" & Chr(13) &
                        "style='font-style: italic;'>" & Date.Now.ToString("yyyyMMdd") & "</span><br>" & Chr(13) &
                        "<big style='font-weight: bold;'><br>" & Chr(13) &
                        "Se Aplico Satisfactoriamente.<br>" & Chr(13) &
                        "<span style='color: rgb(204, 102, 0);'>**Se adjunta archivo de respuesta a sistema de ASG**</span><br>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**Mensaje generado" & Chr(13) &
                        "por el Sistema**</span><br style='color: rgb(0, 102, 0);'>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**No Responder**</span><br>" & Chr(13) &
                        "<br>"
                    Try
                        _SMTP.Send(_Message)
                    Catch ex As SmtpException
                        Logtxtbox.AppendText("Error en conexión SMTP, no se pudo notificar via E-mail." & ex.Message.ToString & vbCrLf)
                    End Try
                    Logtxtbox.AppendText("El archivo : " & nomarchivo & " Fue notificado por correo electronico." & Chr(13))

                Case 2
                    If Len(emailftp) <> 0 And (emailftp) <> "email@email.com" Then
                        _Message.[To].Add(emailftp)
                    End If

                    '_Message.[To].Add(emailnoti1)
                    _Message.[To].Add(emailnotiria)
                    _Message.From = New MailAddress("cobranzas@tarjetasoccidente.hn", "Sistema de Cobranzas - Bancocci", System.Text.Encoding.UTF8)
                    _Message.Subject = "Notificacion"
                    _Message.SubjectEncoding = System.Text.Encoding.UTF8
                    _Message.Priority = System.Net.Mail.MailPriority.High

                    ' ADICION DE DATOS ADJUNTOS
                    Dim _File As String = CarpetaBackup & Date.Now.ToString("yyyyMMdd") & "\" & nombretmp
                    Dim _Attachment As New Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet)
                    _Message.Attachments.Add(_Attachment)

                    _Message.IsBodyHtml = True
                    _Message.Body =
                        "<img style='width: 500px; height: 100px;' alt='LOGO BANCO'" & Chr(13) &
                        "title='LOGO'" & Chr(13) &
                        "src='http://www.bancocci.hn/images/header.png'><span" & Chr(13) &
                        "style='font-weight: bold;'><br>" & Chr(13) &
                        "<br>" & Chr(13) &
                        "El Archivo :&nbsp; </span><span" & Chr(13) &
                        "style='font-style: italic;'>" & nomarchivo & "</span><br>" & Chr(13) &
                        "<big style='font-weight: bold;'><br>" & _Attachment.Name & Chr(13) &
                        "Sistema de renvio mensual.</big><br>" & Chr(13) &
                        "<br>" & Chr(13) &
                        "<span style='text-decoration: underline;'><br>" & Chr(13) &
                        "Envio Satisfactorio.<br>" & Chr(13) &
                        "<span style='color: rgb(204, 102, 0);'>**Se adjunta archivo de respuesta a sistema de ASG**</span><br>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**Mensaje generado" & Chr(13) &
                        "por el Sistema**</span><br style='color: rgb(0, 102, 0);'>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**No Responder**</span><br>" & Chr(13) &
                        "<br>"
                    Try
                        _SMTP.Send(_Message)
                    Catch ex As SmtpException
                        Logtxtbox.AppendText("Error en conexión SMTP, no se pudo notificar via E-mail." & ex.Message.ToString & vbCrLf)
                    End Try
                    Logtxtbox.AppendText("El Archivo : " & nomarchivo & " Fue notificado por correo electronico." & Chr(13))

                Case 3
                    If Len(emailftp) <> 0 And (emailftp) <> "email@email.com" Then
                        _Message.[To].Add(emailftp)
                    End If

                    '_Message.[To].Add(emailnoti1)
                    _Message.[To].Add(emailnotiria)
                    _Message.From = New MailAddress("cobranzas@tarjetasoccidente.hn", "Sistema de Cobranzas - Bancocci", System.Text.Encoding.UTF8)
                    _Message.Subject = "Notificacion"
                    _Message.SubjectEncoding = System.Text.Encoding.UTF8

                    ' ADICION DE DATOS ADJUNTOS
                    Dim _File As String = (CarpetaBackup & Date.Now.ToString("yyyyMMdd") & "\" & nombretmp)
                    Dim _Attachment As New Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet)
                    _Message.Attachments.Add(_Attachment)

                    _Message.IsBodyHtml = True
                    _Message.Body =
                        "<img style='width: 500px; height: 100px;' alt='LOGO BANCO'" & Chr(13) &
                        "title='LOGO'" & Chr(13) &
                        "src='http://www.bancocci.hn/images/header.png'><span" & Chr(13) &
                        "style='font-weight: bold;'><br>" & Chr(13) &
                        "<br>" & Chr(13) &
                        "El Archivo :&nbsp; </span><span" & Chr(13) &
                        "<big style='font-weight: bold;'><br>" & _Attachment.Name & Chr(13) &
                        " - Copia del archivo de salida.</big><br>" & Chr(13) &
                        "<br>" & Chr(13) &
                        "<span style='text-decoration: underline;'><br>" & Chr(13) &
                        "Correo Trasmitido Satisfactoriamente.<br>" & Chr(13) &
                        "<span style='color: rgb(204, 102, 0);'>**Se adjunta archivo de respuesta a sistema de ASG**</span><br>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**Mensaje generado" & Chr(13) &
                        "por el Sistema**</span><br style='color: rgb(0, 102, 0);'>" & Chr(13) &
                        "<span style='color: rgb(0, 102, 0);'>**No Responder**</span><br>" & Chr(13) &
                        "<br>"
                    Try
                        _SMTP.Send(_Message)

                    Catch ex As SmtpException
                        Logtxtbox.AppendText("Error en conexión SMTP, no se pudo notificar via E-mail." & ex.Message.ToString & vbCrLf)
                    End Try
                    Logtxtbox.AppendText("El Archivo : " & nomarchivo & " Fue notificado por correo electronico." & Chr(13))

            End Select
        End If
    End Sub

    ''' <summary>
    ''' Funcion de nombramiento del Log.
    ''' </summary>
    Public Sub tiempolog_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tiempolog.Tick
        If File.Exists(CarpLogs & nomarchs & Date.Now.ToString("dd-MM-yyyy") & ".log") Then
            Logtxtbox.SaveFile(CarpLogs & nomarchs & Date.Now.ToString("dd-MM-yyyy") & ".log", RichTextBoxStreamType.RichText)
        Else
            Logtxtbox.Text = ""
            Logtxtbox.SaveFile(CarpLogs & nomarchs & Date.Now.ToString("dd-MM-yyyy") & ".log", RichTextBoxStreamType.RichText)
        End If
    End Sub

    ''' <summary>
    ''' Botton para entrar a configuracion
    ''' </summary>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim forma As New frmConf
        forma.Show()
    End Sub

    ''' <summary>
    ''' Botton de apertura del log.
    ''' </summary>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim pantallalog As New Logs
        pantallalog.Show()
    End Sub

    ''' <summary>
    ''' Botton del tray de windows.
    ''' </summary>
    Private Sub NotifyIcon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        If Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Minimized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub


    ''' <summary>
    ''' Filtro a traves de fechas para los archivos cargados
    ''' </summary>
    Public Sub Chequeo_Tiempo()
        '10000 = 10 segundos como estaba originalmente
        Dim horaevaluacion, Horamax, ddomingo As String
        horaevaluacion = TimeString
        horaevaluacion = horaevaluacion.Remove(2, Len(horaevaluacion) - 2)

        Horamax = TimeString
        Horamax = Horamax.Remove(5, Len(Horamax) - 5)
        Horamax = Horamax.Remove(2, 1)
        xdia = Today.DayOfWeek.ToString

        If xdia = "Saturday" Then
            horafinalcola = SMaxima
        Else
            horafinalcola = Horamaxima
        End If

        ddomingo = Today.DayOfWeek.ToString

        If Not (ddomingo = "Sunday") Then
            If ((Val(horaevaluacion) >= Val(horainiciocola)) And (Val(horaevaluacion) < Val(horafinalcola))) Then
                Activado = True
            Else
                Activado = False
            End If
        Else
            Activado = False
        End If
    End Sub

    ''' <summary>
    ''' Rebisamos que existan archivos correctos en la cola.
    ''' </summary>
    Public Sub Chequeo_Cola()
        If Not Cola_Nombre_Arch(0) = "" Then
            Logtxtbox.AppendText("Encontre Trabajo en la cola, Procediendo a ejecutar Proceso" & Chr(13))
            proceso(Cola_Direccion_Arch(0), Cola_Nombre_Arch(0))
        End If
    End Sub


    ''' <summary>
    ''' Proceso por el cual subimos archivos al As400 los borramos, sacamos copias y traemos la respuesta.
    ''' </summary>
    Public Sub proceso(ByVal Ruta_Arch_Interno As String, ByVal Nombre_Arch_Interno As String)

        'Creo las listas de Clase ListaArchivos
        Dim ArchsVID As New List(Of ListaArchivos)
        Dim ArchsVIL As New List(Of ListaArchivos)

        Logtxtbox.AppendText("Se ha encontrado trabajo en cola, procediendo a ejecutar PROCESO" & Chr(13))
        Logtxtbox.AppendText("Listando Archivos Bajados..." & Chr(13))

        Try

            ' recorrer los ficheros en la colección
            For Each sFichero As String In Directory.GetFiles(CicloCarp, "*.txt", SearchOption.TopDirectoryOnly)

                ' Crear nuevo objeto FileInfo
                Dim Archivo As New FileInfo(sFichero)

                'Comparo y agrego a la Lista de archivos a procesar
                If Archivo.Name.ToString.Contains("APL") Or Archivo.Name.ToString.Contains("BIL") Then
                    ArchsVIL.Add(New ListaArchivos(Archivo.Name.ToString, Archivo.Name.Substring(7, 8)))
                ElseIf Archivo.Name.ToString.Contains("APD") Or Archivo.Name.ToString.Contains("BID") Then
                    ArchsVID.Add(New ListaArchivos(Archivo.Name.ToString, Archivo.Name.Substring(7, 8)))
                End If

                Logtxtbox.AppendText("Archivo " & Archivo.Name.ToString & " En Lista de Aplicación..." & Chr(13))
            Next

        Catch ex As Exception
            Logtxtbox.AppendText("Existió un Error en el proceso de agrupamiento de dollar y lempira : " & ex.Message & Chr(13))

        End Try


        'Variables de parametros
        Dim p1, p2, p3, p4, p5 As String
        Dim fechaL As Long 'fechaD,

        'Controles de Proceso FTP
        Dim _enviaDolares As New InfoProcesoTarea
        Dim _enviaLempiras As New InfoProcesoTarea
        Dim _BajaArchivoRespuesta As New InfoProcesoTarea

        'Controles de Proceso Carga y Aplicación
        Dim _ejecutaDolares As New InfoProcesoTarea
        Dim _ejecutaBackupDolares As New InfoProcesoTarea
        Dim _ejecutaLempiras As New InfoProcesoTarea
        Dim _ejecutaBackupLempiras As New InfoProcesoTarea
        Dim _aplicaDebitos As New InfoProcesoTarea


        'Cuado para el manejo de las colas.  (que siempre valla dollar y lempira junto)
        'Do While (ArchsVID.Count = 1 And ArchsVIL.Count = 1)

        'FIN: verificamos el preoceso de ingreso dedatos
        Try
            'Validar cuando ambas filas estan lllenas aun
            'If ArchsVID.Count <> 0 And ArchsVIL.Count <> 0 Then
            'Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Se Encontraron datos en cola de Dolares y Lempiras." & vbCrLf)

            'fechaD = Val(ArchsVID(0).Archivo.Fecha)
            fechaL = Val(ArchsVIL(0).Archivo.Fecha)

            'If fechaD = fechaL Then
            'Aplico los dos archivos


            '---------------------------------------------------------------------------------------
            '******************Seccion de Subida de Archivo Dolares al AS400*********************

            'Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & ">>> " & "Subiendo Archivo Dolares... al AS400" & vbCrLf)
            '_enviaDolares = EnviarFTPAS400(CicloCarp & ArchsVID(0).Archivo.Nombre, "/" & "QDLS" & "/" & "BOCAP" & "/" & "PALD" & Date.Now.DayOfYear & ".txt")

            'If _enviaDolares.getExecStatus = True Then
            '    Logtxtbox.AppendText(_enviaDolares.Msg)
            'Else
            '    Logtxtbox.AppendText(_enviaDolares.Msg)
            '    Logtxtbox.AppendText("Error al subir el documento de dollares al AS400")
            '    'Exit Do
            'End If


            ''Se cargan los datos a las tablas temporales del AS400
            'Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Cargando Parametros de aplicación..." & vbCrLf)

            ''Preparando Parametros
            'p2 = "PALD" & Date.Now.DayOfYear & ".TXT"
            'p2 = p2.PadRight(12, " ")

            'p1 = "BOCAP".PadRight(12, " ")
            'p3 = "ASGDAT".PadRight(10, " ")
            'p4 = "PALVID".PadRight(10, " ")
            'p5 = "000"


            'Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Cargando Datos a Tabla Temporal..." & vbCrLf)

            ''Ejecutando Carga
            '_ejecutaDolares = EjecutarCargaArchivoAS400(p1, p2, p3, p4, p5)

            ''Verificando Carga, y realizando Backup
            'If _ejecutaDolares.getExecStatus = True Then
            '    Logtxtbox.AppendText(_ejecutaDolares.Msg)
            'Else
            '    Logtxtbox.AppendText(_ejecutaDolares.Msg)
            '    Logtxtbox.AppendText("Existió un Error en el proceso de Backup el archivo de dollares, transaccion suspendida.")
            '    'Exit Do
            'End If


            '---------------------------------------------------------------------------------------
            '******************Seccion de Subida de Archivo Lempiras al AS400*********************

            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & ">>> " & "Subiendo Archivo Lempiras... al AS400" & vbCrLf)
            _enviaLempiras = EnviarFTPAS400(CicloCarp & ArchsVIL(0).Archivo.Nombre, "/" & "QDLS" & "/" & "BOCAP" & "/" & "BAPL" & Date.Now.DayOfYear & ".txt")

            If _enviaLempiras.getExecStatus = True Then
                Logtxtbox.AppendText(_enviaLempiras.Msg)
            Else
                Logtxtbox.AppendText(_enviaLempiras.Msg)
                Logtxtbox.AppendText("Favor llamar a Informática, Hacer 'Copy y Paste' de este Log y enviarlo por Correo Electrónico al encargado.")
                'Exit Do
            End If

            'Se cargan los datos a las tablas temporales del AS400
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Cargando Parametros de aplicación..." & vbCrLf)

            'Preparando Parametros
            p2 = "BAPL" & Date.Now.DayOfYear & ".TXT"
            p2 = p2.PadRight(12, " ")

            p1 = "BOCAP".PadRight(12, " ")
            p3 = "BAPDAT".PadRight(10, " ")
            p4 = "BAPVIL".PadRight(10, " ")
            p5 = "000"

            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Cargando Datos a Tabla Temporal..." & vbCrLf)

            'Ejecutando Carga
            _ejecutaLempiras = EjecutarCargaArchivoAS400(p1, p2, p3, p4, p5)

            'Verificando Carga, y realizando Backup
            If _ejecutaLempiras.getExecStatus = True Then
                Logtxtbox.AppendText(_ejecutaLempiras.Msg)

            Else
                Logtxtbox.AppendText(_ejecutaLempiras.Msg)
                Logtxtbox.AppendText("Existió un Error en el proceso de Backup el archivo de lempiras, transaccion suspendida.")
                'Exit Do
            End If

            '---------------------------------------------------------------------------------------
            '******************Ejecutamos la aplicacion a las cuentas*********************

            If _ejecutaLempiras.getExecStatus = True Then ' And _ejecutaLempiras.getExecStatus = True Then
                'Ejecuta Aplicación

                _aplicaDebitos = EjecutaAplicacionAS400()
                If _aplicaDebitos.getExecStatus = True Then
                    Logtxtbox.AppendText(_aplicaDebitos.Msg)

                    '---------------------------------------------------------------------------------------
                    '******************Transaccion Exitosa*********************

                    Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "---------------------------------------------------" & vbCrLf)
                    'Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo " & ArchsVID(0).Archivo.Nombre & " Aplicado Correctamente..." & vbCrLf)
                    Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo " & ArchsVIL(0).Archivo.Nombre & " Aplicado Correctamente..." & vbCrLf)
                    Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "---------------------------------------------------" & vbCrLf)

                    'Se bajan los reportes a la carpeta de salida. se prepara los datos de salida
                    BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "BAPFRE.TXT", CarpetaBackup & "RBOAPL" & Date.Now.ToString("yyyyMMdd") & "1.TXT")
                    PrepararSalidaTemp("RBOAPL" & Date.Now.ToString("yyyyMMdd") & "1.TXT")

                    'BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "PALFRD.TXT", CarpetaBackup & "RBOAPD" & Date.Now.ToString("yyyyMMdd") & "1.TXT")
                    'PrepararSalidaTemp("RBOAPD" & Date.Now.ToString("yyyyMMdd") & "1.TXT")

                    'Se ejecuta el backup del ultimo ciclo 'Prevencion de perdida de conexion con el AS400
                    _ejecutaBackupLempiras = BackupArchivo(CicloCarp & ArchsVIL(0).Archivo.Nombre, ArchsVIL(0).Archivo.Nombre & "Cicle")
                    '_ejecutaBackupDolares = BackupArchivo(CicloCarp & ArchsVID(0).Archivo.Nombre, ArchsVID(0).Archivo.Nombre & "Cicle")

                    'Se bajan los datos para el siguiente ciclo
                    'BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "PALVID.TXT", CicloCarp & ArchsVID(0).Archivo.Nombre)
                    BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "BAPVIL.TXT", CicloCarp & ArchsVIL(0).Archivo.Nombre)

                    'Libero Posiciones.
                    'ArchsVID.Remove(ArchsVID.Item(0))
                    ArchsVIL.Remove(ArchsVIL.Item(0))

                Else
                    Logtxtbox.AppendText(_aplicaDebitos.Msg)
                    Logtxtbox.AppendText(">>> ERROR al aplicar los debitos a las cuentas")
                    'Exit Do
                End If
            End If


            '---------------------------------------------------------------------------------------
            '///////////////////Procedimiento cuando no es igual la fecha de los archivos///////////////////

            ' Else 'Fechas tienen que ser iguales
            'Logtxtbox.AppendText(_aplicaDebitos.Msg)
            ' Logtxtbox.AppendText(">>> ERROR al aplicar los debitos a las cuentas, Fechas no son iguales")
            'Exit Do

            'End If

            'Else 'Un archivo de dollar u lempira
            '    'Logtxtbox.AppendText(_aplicaDebitos.Msg)
            '    Logtxtbox.AppendText(">>> ERROR al aplicar los debitos a las cuentas, Se nesecita un archivo de dollares y lempiras")
            '    'Exit Do
            'End If

        Catch ex As Exception
            Logtxtbox.AppendText("No se puede procesar la cola, Se encontro un error en el proceso de carga de archivos..." & ex.Message.ToString & Chr(13))
        End Try

        'Loop

        'If ArchsVID.Count = 1 And ArchsVIL.Count > 1 Then
        '    Logtxtbox.AppendText("---No se encontro archivo en dollares, terminando ciclo***.---" & Chr(13))
        '    _ejecutaBackupLempiras = BackupArchivo(CicloCarp & ArchsVIL(0).Archivo.Nombre, ArchsVIL(0).Archivo.Nombre)
        'End If

        'If ArchsVID.Count > 1 And ArchsVIL.Count = 1 Then
        '    Logtxtbox.AppendText("---No se encontro archivo en Lempiras, terminando ciclo***.---" & Chr(13))
        '    _ejecutaBackupDolares = BackupArchivo(CicloCarp & ArchsVID(0).Archivo.Nombre, ArchsVID(0).Archivo.Nombre)
        'End If

        'Reinicio variable
        Cola_Nombre_Arch(0) = ""
        Logtxtbox.AppendText("---Listo para más trabajo.---" & Chr(13))

    End Sub

    ''' <summary>
    ''' Envia por FTP al AS400
    ''' </summary>
    ''' <param name="_Local">The _ local.</param>
    ''' <param name="_Remoto">The _ remoto.</param>
    ''' 
    Public Function EnviarFTPAS400(ByVal _Local As String, ByVal _Remoto As String) As InfoProcesoTarea
        Dim _out As New InfoProcesoTarea
        Dim _ftp As New Rebex.Net.Ftp
        'Provado
        Try
            _ftp.Connect(dirftp)
            _ftp.Login(usuario, password)
            _ftp.PutFile(_Local, _Remoto)
            _ftp.Disconnect()
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo enviado Correctamente..." & vbCrLf
            _out.getExecStatus = True
        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al Enviar Archivo: " & ex.Message & vbCrLf
            _out.getExecStatus = False
        End Try
        _ftp = Nothing

        EnviarFTPAS400 = _out



        'Dim _out As New InfoProcesoTarea
        'Dim _ftp As New Rebex.Net.Ftp
        ''Provado
        'Try
        '    _ftp.Connect(dirftp)
        '    _ftp.Login(usuario, password)
        '    _ftp.PutFile(_Local, _Remoto)
        '    _ftp.Disconnect()
        '    _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo enviado Correctamente..." & vbCrLf
        '    _out.getExecStatus = True
        'Catch ex As Exception
        '    _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al Enviar Archivo: " & ex.Message & vbCrLf
        '    _out.getExecStatus = False
        'End Try
        '_ftp = Nothing

        'EnviarFTPAS400 = _out
    End Function

    ''' <summary>
    ''' Bajar Archivo FTP
    ''' </summary>
    ''' <param name="_Remoto">The _ remoto.</param>
    ''' <param name="_Local">The _ local.</param>
    ''' <returns></returns>
    Public Function BajarFTPAS400(ByVal _Remoto As String, ByVal _Local As String) As InfoProcesoTarea
        Dim _out As New InfoProcesoTarea
        Dim _ftp As New Rebex.Net.Ftp

        Try
            _ftp.Connect(dirftp)
            _ftp.Login(usuario, password)
            _ftp.GetFile(_Remoto, _Local)
            _ftp.Disconnect()
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Reporte-AS Bajado Correctamente..." & _Local & vbCrLf)
            _out.getExecStatus = True
        Catch ex As Exception
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error en el Proceso de Bajada en AS400... " & _Local & ex.Message & vbCrLf)
            _out.getExecStatus = False
        End Try
        _ftp = Nothing

        BajarFTPAS400 = _out
    End Function

    ''' <summary>
    ''' Ejecuta la carga archivo A S400.
    ''' </summary>
    ''' <param name="p1">The p1.</param>
    ''' <param name="p2">The p2.</param>
    ''' <param name="p3">The p3.</param>
    ''' <param name="p4">The p4.</param>
    ''' <param name="p5">The p5.</param>
    ''' <returns></returns>
    Public Function EjecutarCargaArchivoAS400(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String) As InfoProcesoTarea
        Dim CON1 As New ADODB.Connection
        Dim Cmd1 As New ADODB.Command
        Dim _out As New InfoProcesoTarea

        Try

            vDireccioIpAs400 = dirftp
            vUsuarioAs400 = usuario
            vContraseñaAs400 = password

            CON1.Open("Provider=IBMDA400;Data Source=" & vDireccioIpAs400 & ";", vUsuarioAs400, vContraseñaAs400)

        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al Abrir conexion con el AS400 : " & ex.Message.ToString & vbCrLf
            _out.getExecStatus = False
            EjecutarCargaArchivoAS400 = _out
            Exit Function
        End Try

        Try
            Cmd1.ActiveConnection = CON1
            Cmd1.CommandTimeout = 25
            Cmd1.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd1.CommandText = "{{CALL /QSYS.LIB/BAPPGM.LIB/BALOAD.PGM(?,?,?,?,?)}}"
            Cmd1.Parameters.Append(Cmd1.CreateParameter("A", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 12, p1))
            Cmd1.Parameters.Append(Cmd1.CreateParameter("B", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 12, p2))
            Cmd1.Parameters.Append(Cmd1.CreateParameter("C", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 10, p3))
            Cmd1.Parameters.Append(Cmd1.CreateParameter("D", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 10, p4))
            Cmd1.Parameters.Append(Cmd1.CreateParameter("E", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 3, p5))

            Cmd1.Prepared = True
            Cmd1.Execute()

            If Cmd1.Parameters(4).Value <> "145" Then
                'ERROR
                _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error en el Proceso de Carga en AS400... " & vbCrLf
                _out.getExecStatus = False
            ElseIf Cmd1.Parameters(4).Value = "145" Then
                'CORRECTO
                _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo Cargado... " & vbCrLf
                _out.getExecStatus = True
            End If
            Cmd1.Cancel()
            CON1.Close()
        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al ejecutar carga de archivo temporales en AS400:  " & ex.Message & vbCrLf
            _out.getExecStatus = False
            Cmd1.Cancel()
            CON1.Close()
        End Try

        Cmd1 = Nothing
        CON1 = Nothing

        EjecutarCargaArchivoAS400 = _out
    End Function

    ''' <summary>
    ''' Backups el archivo.
    ''' </summary>
    Public Function BackupArchivo(ByVal _rutaOrigen As String, ByVal _nombreBackup As String) As InfoProcesoTarea
        Dim _out As New InfoProcesoTarea
        Try
            ' Sub carpetas
            Dim dir As New System.IO.DirectoryInfo(CarpetaBackup & Date.Now.ToString("yyyyMMdd"))
            If Not dir.Exists Then
                Directory.CreateDirectory(CarpetaBackup & Date.Now.ToString("yyyyMMdd"))
            End If

            FileCopy(_rutaOrigen, CarpetaBackup & Date.Now.ToString("yyyyMMdd") & "\" & _nombreBackup)

            File.Delete(_rutaOrigen)
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo Copiado a Backup Satisfactoriamente" & vbCrLf
            _out.getExecStatus = True

        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al realizar el backup:  " & ex.Message & vbCrLf
            _out.getExecStatus = False
        End Try

        BackupArchivo = _out
    End Function

    ''' <summary>
    ''' ********IMPORTANTE******* Ejecuta la Aplicación de los datos en el AS400 sobre las tablas temporales
    ''' </summary>
    ''' <returns></returns>
    Public Function EjecutaAplicacionAS400() As InfoProcesoTarea
        Dim CON2 As New ADODB.Connection
        Dim Cmd2 As New ADODB.Command
        Dim _out As New InfoProcesoTarea

        Dim parm1 As String
        parm1 = "000"

        Try
            vDireccioIpAs400 = dirftp
            vUsuarioAs400 = usuario
            vContraseñaAs400 = password

            CON2.Open("Provider=IBMDA400;Data Source=" & vDireccioIpAs400 & ";", vUsuarioAs400, vContraseñaAs400)

        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error al Abrir conexion con el AS400 : " & ex.Message.ToString & vbCrLf
            _out.getExecStatus = False
            EjecutaAplicacionAS400 = _out
            Exit Function
        End Try

        Try
            Cmd2.CommandTimeout = 25
            Cmd2.ActiveConnection = CON2
            Cmd2.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd2.CommandText = "{{CALL /QSYS.LIB/BAPPGM.LIB/BAACDC01.PGM(?)}}"
            Cmd2.Parameters.Append(Cmd2.CreateParameter("A", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 3, parm1))
            Cmd2.Prepared = True

            Cmd2.Execute()

            If Cmd2.Parameters(0).Value <> "145" Then
                'ERROR
                _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error en el Proceso de Aplicación en AS400... " & vbCrLf
                _out.getExecStatus = False
            ElseIf Cmd2.Parameters(0).Value = "145" Then
                'CORRECTO
                _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo Aplicado Correctamente... " & vbCrLf
                _out.getExecStatus = True
            End If
            Cmd2.Cancel()
            CON2.Close()
        Catch ex As Exception
            _out.Msg = Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error :  " & ex.Message & vbCrLf
            _out.getExecStatus = False
            Cmd2.Cancel()
            CON2.Close()
        End Try

        Cmd2 = Nothing
        CON2 = Nothing

        EjecutaAplicacionAS400 = _out
    End Function

    ''' <summary>
    ''' Clase que permite manejar Lista de Archivos.
    ''' </summary>
    Public Class ListaArchivos
        Public Structure Archivos
            Dim Nombre As String
            Dim Fecha As String
        End Structure
        Public Archivo As New Archivos
        Public Sub New(ByVal _inNombre As String, ByVal _inFecha As String)
            Archivo.Nombre = _inNombre
            Archivo.Fecha = _inFecha.Substring(0, 8)
        End Sub
    End Class

    Private Sub BManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'SubirSFTPASG(Cola_Direccion_Arch(0), Cola_Nombre_Arch(0))
        BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "BAPFRE.TXT", Manual & "ROCHBIL" & Date.Now.ToString("yyyyMMdd") & "1.TXT")
        'BajarFTPAS400("/" & "QDLS" & "/" & "BOCAP" & "/" & "PALFRD.TXT", Manual & "ROCHBID" & Date.Now.ToString("yyyyMMdd") & "1.TXT")
        Logtxtbox.AppendText("Descarga manual finalizado, esperando archivos")
    End Sub

    ''' <summary>
    ''' Bajar Archivo SFTP desde el servidor de aplicacion
    ''' </summary>
    ''' <param name="_Remoto">The _ remoto.</param>
    ''' <param name="_Local">The _ local.</param>
    ''' <returns></returns>
    ''' 
    Public Function BajarSFTPASG(ByVal _Remoto As String, ByVal _Local As String) As InfoProcesoTarea

        ' Creo una instancia de la clase FTP 
        Dim _Sftp As New SftpClient(ASGIp, ASGUser, ASGPass)

        Dim _out As New InfoProcesoTarea
        Dim CarpetaOutBox As String = "/desdeAFFINITY"

        Try
            ' Create an SFTP client
            Using Sftp As New SftpClient(ASGIp, ASGUser, ASGPass)
                ' Connect to the SFTP server
                Sftp.Connect()

                ' Retrieve and display the list of files and directories
                Dim list As IEnumerable(Of ISftpFile) = Sftp.ListDirectory(CarpetaOutBox)

                For Each file In list

                    ' Download the file from the SFTP server to the local machine
                    'Using localFileStream As New FileStream(CicloCarp & "/" & Path.GetFileName(file.FullName), FileMode.Create)
                    '    Sftp.DownloadFile(CarpetaOutBox & "/" & Path.GetFileName(file.FullName), localFileStream)
                    'End Using


                    ' Skip directories
                    If file.IsDirectory Then
                        Continue For
                    End If

                    ' Build local and remote file paths
                    Dim localFilePath As String = Path.Combine(Entradas, file.Name)
                    Dim remoteFilePath As String = Path.Combine(CarpetaOutBox, file.Name)

                    ' Download the file from the SFTP server to the local machine
                    Using localFileStream As New FileStream(localFilePath, FileMode.Create)
                        Sftp.DownloadFile(remoteFilePath, localFileStream)
                    End Using


                    Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo-ASG Descargado Correctamente..." & "--" & vbCrLf)
                    _out.getExecStatus = True

                    ' Delete the file on the server
                    Sftp.DeleteFile(file.FullName)
                    nombretmp = file.Name


                    ' Check if it's a text file
                    If file.Name.EndsWith(".TXT", StringComparison.OrdinalIgnoreCase) Then
                        ' Prepare the file for further processing and delete it from the local directory
                        Preparar(nombretmp)
                    Else
                        ' If it's a monthly file
                        BackupArchivo(Path.Combine(Entradas, nombretmp), nombretmp)
                        EnvioMail(2)
                    End If
                Next

                ' Disconnect from the SFTP server
                Sftp.Disconnect()
            End Using

        Catch ex As Exception
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error en el Proceso de Descarga de archivo ASG... " & ex.Message & "--" & vbCrLf)
                _out.getExecStatus = False
            End Try

            BajarSFTPASG = _out
    End Function

    Sub SaveStreamToFile(stream As Stream, filePath As String)
        ' Ensure the stream is at the beginning
        stream.Seek(0, SeekOrigin.Begin)

        ' Create a StreamWriter to write to the file
        Using writer As New StreamWriter(filePath)
            ' Copy the stream to the StreamWriter
            Using reader As New StreamReader(stream)
                writer.Write(reader.ReadToEnd())
            End Using
        End Using
    End Sub


    ''' <summary>
    ''' Subir Archivo SFTP hasta el servidor de aplicacion
    ''' </summary>
    ''' <param name="_Remoto">The _ remoto.</param>
    ''' <param name="_Local">The _ local.</param>
    ''' <returns></returns>
    Public Function SubirSFTPASG(ByVal _Remoto As String, ByVal _Local As String) As InfoProcesoTarea

        Dim _Sftp As New SftpClient(ASGIp, ASGUser, ASGPass)

        Dim _out As New InfoProcesoTarea
        Dim CarpetaInBox As String = "/paraAFFINITY"

        Try
            ' Connect to the SFTP server
            _Sftp.Connect()

            ' Upload files from the specified local directory to the SFTP server
            For Each fileName As String In Directory.GetFiles(Salidas)
                Dim remoteFilePath As String = CarpetaInBox & "/" & Path.GetFileName(fileName)

                Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Se inicia la carga del Archivo al SFTP " & vbCrLf)

                Using fileStream As FileStream = File.OpenRead(fileName)
                    _Sftp.UploadFile(fileStream, remoteFilePath)
                End Using

                Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Se Inicia el borrado de la carpeta Salida y se coloca en:" & CarpetaBackup & Date.Now.ToString("yyyyMMdd") & "\" & vbCrLf)

                ' Backup the file (move to the backup directory)
                BackupArchivo(fileName, Path.GetFileName(fileName))

                Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Se Inicia el envio por correo a la persona configurada en emailftp" & vbCrLf)

                EnvioMail(3)

                Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Envio Exitosa de la data--" & vbCrLf)
            Next

            ' Disconnect from the SFTP server
            _Sftp.Disconnect()

            ' Email de salida
            EnvioMail(1)

        Catch ex As Exception
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Error en el Proceso de Subida de archivo ASG... " & ex.Message & "--" & vbCrLf)
            _out.getExecStatus = False

        Finally
            Logtxtbox.AppendText(Date.Now.ToString("dd/MM/yy HH:mm:ss ") & "> " & "Archivo-ASG Proceso de salida terminado" & "--" & vbCrLf)
            _out.getExecStatus = True
        End Try

        _Sftp = Nothing
        'Email de salida
        EnvioMail(1)

    End Function

    ''' <summary>
    ''' Prepara los documentos despues de mandarlos a traer y los coloca en ciclo, ademas borra de entrada
    ''' </summary>
    ''' <returns></returns>
    Public Function Preparar(ByVal Name As String) As Boolean

        Dim objReader As New StreamReader(Entradas & Name)
        Dim sLine As String = ""
        Dim arrText As New ArrayList()
        Dim TextFilePreparado As New StreamWriter(CicloCarp & Name)

        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                arrText.Add(sLine)
            End If
        Loop Until sLine Is Nothing
        objReader.Close()

        For Each sLine In arrText
            sLine = sLine.Substring(0, sLine.Length)
            sLine = sLine + "0"

            'Escibo cada linea en el documento prosesado
            TextFilePreparado.WriteLine(sLine)

        Next
        TextFilePreparado.Close()

        'Borro el archivo de la carpeta de entrada
        BackupArchivo(Entradas & Name, Name)
    End Function

    ''' <summary>
    ''' Prepara los documentos temporales de cada ciclo
    ''' </summary>
    ''' <returns></returns>
    Public Function PrepararSalidaTemp(ByVal Name As String) As Boolean

        Dim objReader As New StreamReader(CarpetaBackup & Name)
        Dim sLine As String = ""
        Dim arrText As New ArrayList()
        Dim TextFilePreparado As New StreamWriter(Salidas & Name, True)

        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                arrText.Add(sLine)
            End If
        Loop Until sLine Is Nothing
        objReader.Close()

        For Each sLine In arrText
            If sLine.Contains("Cobro Realizado") Then
                'Escibo cada linea en el documento prosesado
                TextFilePreparado.WriteLine(sLine)
            End If
        Next

        TextFilePreparado.Close()

    End Function

    ''' <summary>
    ''' Prepara los documentos de salida
    ''' </summary>
    ''' <returns></returns>
    Public Function PrepararSalida(ByVal Name As String) As Boolean
        If Name.Contains("APD") Then
            Dim objReader As New StreamReader(CarpetaBackup & Name)
            Dim sLine As String = ""
            Dim arrText As New ArrayList()
            Dim TextFilePreparado As New StreamWriter(Salidas & Name, True)

            Do
                sLine = objReader.ReadLine()
                If Not sLine Is Nothing Then
                    arrText.Add(sLine)
                End If
            Loop Until sLine Is Nothing
            objReader.Close()

            For Each sLine In arrText
                If Not sLine.Contains("Cobro Realizado") Then
                    'Escibo cada linea en el documento prosesado
                    TextFilePreparado.WriteLine(sLine)
                End If
            Next

            TextFilePreparado.Close()
            'Borro el archivo temporal 
            My.Computer.FileSystem.DeleteFile(CarpetaBackup & Name)

        End If

        If Name.Contains("APL") Then
            Dim objReader As New StreamReader(CarpetaBackup & Name)
            Dim sLine As String = ""
            Dim arrText As New ArrayList()
            Dim TextFilePreparado As New StreamWriter(Salidas & Name, True)

            Do
                sLine = objReader.ReadLine()
                If Not sLine Is Nothing Then
                    arrText.Add(sLine)
                End If
            Loop Until sLine Is Nothing
            objReader.Close()

            For Each sLine In arrText
                If Not sLine.Contains("Cobro Realizado") Then
                    'Escibo cada linea en el documento prosesado
                    TextFilePreparado.WriteLine(sLine)
                End If
            Next

            TextFilePreparado.Close()
            'Borro el archivo temporal 
            My.Computer.FileSystem.DeleteFile(CarpetaBackup & Name)
        End If

    End Function

    Private Sub Ciclo_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ciclo.Tick

        'Cada hora es = 3 600 000
        Dim Directorio As New System.IO.DirectoryInfo(CicloCarp)
        Dim DirFiles As System.IO.FileInfo() = Directorio.GetFiles()

        Dim DirectorioSalidas As New System.IO.DirectoryInfo(Salidas)
        Dim DirFilesSalidas As System.IO.FileInfo() = DirectorioSalidas.GetFiles()

        Dim DirectorioSalidasTemp As New System.IO.DirectoryInfo(CarpetaBackup)
        Dim DirFilesSalidasTemp As System.IO.FileInfo() = DirectorioSalidasTemp.GetFiles()

        Dim File As System.IO.FileInfo
        Dim SalidaFlag As Boolean = False

        'list the names of all files in the specified directory
        Try
            'Reviso si es hora de trabajo
            Chequeo_Tiempo()

            If Activado = True Then
                If Directorio.GetFiles.GetLength(0) = 0 Then
                    'Bajo los archivos del servidor SFTP y los formateo para su uso
                    BajarSFTPASG(Cola_Direccion_Arch(0), Cola_Nombre_Arch(0))
                End If

                'Muevo los documentos de salida del ciclo
                For Each File In DirFiles
                    If File.Name.Contains("FRE") Or File.Name.Contains("FRD") Then
                        My.Computer.FileSystem.MoveFile(File.Name, Salidas)
                    Else
                        rarchivo = CicloCarp & File.Name
                        nomarchivo = File.Name

                        'Logtxtbox.AppendText("Trabajando con achivo: " & nomarchivo & ". -->" & Date.Now.Hour.ToString & ":" & Date.Now.Minute.ToString & Chr(13))
                        Dim Contador_Colas As Integer = 0
                        Dim Ultima_Pos_Cola As Integer = 0

                        While Cola_Nombre_Arch(Contador_Colas) <> ""
                            Contador_Colas += 1
                        End While

                        Ultima_Pos_Cola = Contador_Colas
                        Logtxtbox.AppendText("Ingresando archivos a la cola de trabajo." & Chr(13))
                        Cola_Nombre_Arch(Ultima_Pos_Cola) = nomarchivo
                        Cola_Direccion_Arch(Ultima_Pos_Cola) = rarchivo

                        'Reviso la cola existente y sea la correcta
                        Chequeo_Cola()
                        Exit For

                    End If
                Next
            Else 'Si es fuera de hora de trabajo


                'Borro los datos del ciclo temporal
                For Each File In DirFiles
                    BackupArchivo(CicloCarp & File.Name, File.Name & "Final.txt")
                Next

                'Armo el archivo de salida final
                For Each File In DirFilesSalidasTemp
                    PrepararSalida(File.Name)
                Next

                'Ingreso los datos para la salida de informacion y me aseguro que exista
                For Each File In DirFilesSalidas

                    SubirSFTPASG(Cola_Direccion_Arch(0), Cola_Nombre_Arch(0))

                    'Salir del for si no hay datos dentro del primero (solo nesecito uno)
                    Exit For
                Next
            End If 'If Activado

        Catch ex As Exception
            Logtxtbox.AppendText("Error al accesar archivo...reintentando..." & ex.Message)

        End Try

    End Sub

End Class
