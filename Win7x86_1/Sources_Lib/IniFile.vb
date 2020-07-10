Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Ini

    Public Class IniFile

#Region "ИМПОРТ DLL"

        ''' <summary>
        ''' Записывает ключ в заданный раздел INI-файла.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        ''' <param name="value">Значение ключа.</param>
        ''' <param name="filePath">Путь к INI-файлу.</param>
        <DllImport("kernel32")>
        Private Shared Function WritePrivateProfileString(ByVal section As String, ByVal key As String, ByVal value As String, ByVal filePath As String) As Long
        End Function

        ''' <summary>
        ''' Считывает ключ заданного раздела INI-файла.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        ''' <param name="[default]"></param>
        ''' <param name="retVal"></param>
        ''' <param name="size"></param>
        ''' <param name="filePath">Путь к INI-файлу.</param>
        ''' <remarks>С помощью конструктора записываем путь до файла и его имя. </remarks>
        <DllImport("kernel32")>
        Private Shared Function GetPrivateProfileString(ByVal section As String, ByVal key As String, ByVal [default] As String, ByVal retVal As StringBuilder, ByVal size As Integer, ByVal filePath As String) As Integer
        End Function

#End Region '/ИМПОРТ DLL

#Region "КОНСТРУКТОР"

        ''' <summary>
        ''' Имя файла.
        ''' </summary>
        Private IniPath As String

        ''' <summary>
        ''' Читаем ini-файл и возвращаем значение указного ключа из заданной секции. 
        ''' </summary>
        ''' <param name="iniPath"></param>
        Public Sub New(ByVal iniPath As String)
            Me.IniPath = New FileInfo(iniPath).FullName.ToString
        End Sub

        Public Sub New()
            Me.IniPath = ""
        End Sub

        Public Sub SetIniPath(ByVal iniPath As String)
            Me.IniPath = New FileInfo(iniPath).FullName.ToString
        End Sub

#End Region '/КОНСТРУКТОР

#Region "МЕТОДЫ РАБОТЫ С INI-ФАЙЛАМИ"

        ''' <summary>
        ''' Проверяет, что заданный ключ существует в INI-файле.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        Public Function KeyExists(ByVal section As String, ByVal key As String) As Boolean
            Return (Me.ReadKey(section, key).Length > 0)
        End Function

        ''' <summary>
        ''' Читает значение заданного ключа в заданном разделе INI-файла.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        Public Function ReadKey(ByVal section As String, ByVal key As String) As String
            Dim retVal As New StringBuilder(&HFF)
            IniFile.GetPrivateProfileString(section, key, "", retVal, &HFF, Me.IniPath)
            Return retVal.ToString()
        End Function

        ''' <summary>
        ''' Создаёт заданный ключ в заданном разделе. Если раздел не существует, он будет создан.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        ''' <param name="value">Значение ключа. Если NULL, то ключ будет удалён. Если String.Empty, то присвоится пустое значение.</param>
        Public Sub WriteKey(ByVal section As String, ByVal key As String, ByVal value As String)
            IniFile.WritePrivateProfileString(section, key, value, Me.IniPath)
        End Sub

        ''' <summary>
        ''' Удаляет заданный ключ из заданного раздела INI-файла.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        ''' <param name="key">Имя ключа.</param>
        Public Sub DeleteKey(ByVal section As String, ByVal key As String)
            Me.WriteKey(section, key, Nothing)
        End Sub

        ''' <summary>
        ''' Удаляет заданный раздел INI-файла.
        ''' </summary>
        ''' <param name="section">Имя раздела.</param>
        Public Sub DeleteSection(ByVal section As String)
            Me.WriteKey(section, Nothing, Nothing)
        End Sub

#End Region '/МЕТОДЫ РАБОТЫ С INI-ФАЙЛАМИ

    End Class

End Namespace
