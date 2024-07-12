Imports System.IO

Public Class Logs
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' agregar columnas
        Dim carplogs As String

        Dim sh As Object
        sh = CreateObject("WScript.Shell")
        carplogs = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\CarpLogs")
        frmPrincipalASG.CicloCarp = sh.RegRead("HKEY_CURRENT_USER\Software\BancocciASG\PathCarpeta")
        sh = Nothing

        With ListaArchivos
            .Columns.Add("Nombre", 150)
            .Columns.Add("Fecha y hora de modificación", 150)
            .Columns.Add("Tamaño - bytes ", 100)
            .Columns.Add("Extensión", 80)
            .Columns.Add("Path", 150)
            .View = View.Details
            .GridLines = True
        End With

        ListaArchivos.Items.Clear()
        Try
            ' recorrer los ficheros en la colección
            For Each sFichero As String In Directory.GetFiles( _
                                           carplogs, "*.log", _
                                           SearchOption.TopDirectoryOnly)

                ' Crear nuevo objeto FileInfo
                Dim Archivo As New FileInfo(sFichero)
                ' Crear nuevo objeto ListViewItem
                Dim item As New ListViewItem(Archivo.Name.ToString)
                ' cargar los datos y las propiedades
                With item

                    ' LastWriteTime - fecha de modificación
                    .SubItems.Add(Archivo.LastWriteTime.ToShortDateString & " " & _
                                  Archivo.LastWriteTime.ToShortTimeString)
                    ' Length - tamaño en bytes
                    .SubItems.Add(Archivo.Length.ToString)
                    ' Extension - extensión  
                    .SubItems.Add(Archivo.Extension.ToString)
                    ' Path
                    .SubItems.Add(Archivo.FullName.ToString)
                    ListaArchivos.Items.Add(item) ' añadir el item 
                End With

            Next
            ' errores
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
            Beep()
        End Try

    End Sub

    Private Sub ListaArchivos_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListaArchivos.SelectedIndexChanged
        Dim I As Integer
        For I = 0 To ListaArchivos.Items.Count - 1
            If ListaArchivos.Items(I).Selected Then
                LogtxtBoxReader.LoadFile(ListaArchivos.Items(I).SubItems(4).Text.ToString(), RichTextBoxStreamType.RichText)
            End If
        Next
    End Sub
End Class