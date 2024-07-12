Module codigo
    Public Enum tipoarchivos
        Ria = 1
        InterMex = 2
        ViAmerica = 3
        Reset = 4
    End Enum
    Public Enum tipocorreo
        Correcto = 1
        Fallido = 2
        ProTrans = 3
        Manual = 4
    End Enum
    Public pathnuevoarch As String
    Public rarchivo, nomarchivo, tipoarchivo, email, direccion, stringcorreo, carpeta, usuario, password, dirftp As String
    Public NumRia, Horamaxima, ftplib, Horainiciar, HInicio, SMaxima As String
    ' inicializar el array para los encabezados de columna
    Public aCol() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"}
    Public aCol1() As String = {"1", "2", "3", "4", "5", "6"}
    Public contador As Integer
    Public contador1 As Integer
    Public contador2 As Integer
    Public contador3 As Integer
    Public contador4 As Integer
    Public CarpProcesados, CarpErr, CarpManual, CarpSal As String
    Public ftplibdown, ftpcarpdown, NomArchAplic, NomArchCance, NomArchDiario As String
    Public horaevaluacion, carpetaformateados, confnomarchivo, horainicio As String
    Public nombretmp, nomdownload As String

End Module
