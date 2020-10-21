Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.StringBuilder

Namespace DSStatusyFedExWS.Loggers
#Region "ServiceLogger"

    ''' <summary>
    ''' Klasa odpowiedzialna za zapis logów usługi.
    ''' </summary>
    Public Class ServiceLogger
#Region "Zmienne prywatne"

        Private _logsFilesPath As String = ""

#End Region

#Region "Konstruktory"

        ''' <summary>
        ''' Tworzy instancję klasy ServiceLogger. 
        ''' </summary>
        Public Sub New()
            JoinLines = True
        End Sub

#End Region

#Region "Metody publiczne"

        ''' <summary>
        ''' Dopisuje podany tekst do pliku loga.
        ''' </summary>
        Public Overridable Sub Log(log__1 As String, section As String, twoSections As Boolean)
            Try

                Dim logFile As String = LogFilePath
                Directory.CreateDirectory(Path.GetDirectoryName(logFile))
                section = If(IsTextAfterTrim(section), " " & section & " ", " ")
                Dim s As String = (DateTime.Now & section & log__1 & (If(twoSections, section, ""))).Trim()

                Using sm As New StreamWriter(logFile, True, UTF8Encoding.UTF8)
                    If JoinLines Then
                        s = s.Replace(vbLf, " ")
                    End If
                    sm.WriteLine(s)
                End Using

            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Dopisuje podany tekst do pliku loga jako info.
        ''' </summary>
        Public Overridable Sub LogInfo(log__1 As String)
            If log__1.Trim().Length > 0 Then
                Log(log__1, "INFO", False)
            End If
        End Sub

        ''' <summary>
        ''' Dopisuje podany tekst do pliku loga jako błąd.
        ''' </summary>
        Public Overridable Sub LogError(log__1 As String)
            If log__1.Trim().Length > 0 Then
                Log(log__1, "ERROR", False)
            End If
        End Sub

        ''' <summary>
        ''' Dopisuje do pliku loga listę przekazanych błędów.
        ''' </summary>
        Public Overridable Sub LogErrors(logs As List(Of String))
            For Each log As String In logs
                LogError(log)
            Next
        End Sub

#End Region

#Region "Właściwości"

        ''' <summary>
        ''' Czy linie w logu mają być złączone.
        ''' </summary>
        Public Property JoinLines() As Boolean
            Get
                Return m_JoinLines
            End Get
            Set(value As Boolean)
                m_JoinLines = Value
            End Set
        End Property
        Private m_JoinLines As Boolean

        ''' <summary>
        ''' Katalog logów usługi.
        ''' </summary>
        Public Property LogsDir() As String
            Get
                Dim s As String = _logsFilesPath
                If String.IsNullOrEmpty(s) Then
                    s = ""
                End If
                Dim i As Integer = s.Trim().Length
                If i > 0 Then
                    If Not s(i - 1).ToString().Equals("\") Then
                        s += "\"
                    End If
                    _logsFilesPath = s
                Else
                    Dim a As Assembly = Assembly.GetEntryAssembly()
                    _logsFilesPath = Path.GetDirectoryName(a.Location) & "\Logs\"
                End If
                Return _logsFilesPath
            End Get
            Set(value As String)
                _logsFilesPath = value
            End Set
        End Property

        ''' <summary>
        ''' Ścieżka do aktualnego pliku logów usługi.
        ''' </summary>
        Public ReadOnly Property LogFilePath() As String
            Get
                Dim now As DateTime = DateTime.Now
                Dim a As Assembly = Assembly.GetEntryAssembly()
                Return LogsDir & now.Year & "\" & now.Month & "\" & now.Day & " " & a.GetName().Name & ".log"
            End Get
        End Property


        Public Function LogXMLFileP(projektNazwa As String, xmlFileId As String, xmlFileText As String) As Integer

            Dim retVal As Integer = 0
            Dim LogFileDir = "SuperPakerLogs"

            If xmlFileText IsNot Nothing Then

                Dim datePath As String = ""
                Dim fileLogDirectory As String = ""
                Dim Year As String = DateTime.Now.Year.ToString()
                Dim Month As String = DateTime.Now.Month.ToString()
                Dim Day As String = DateTime.Now.Day.ToString()

                datePath = Year & "/" & Month & "/" & Day
                'System.Web.HttpContext.Current.Server.MapPath("/")
                fileLogDirectory = System.Web.HttpContext.Current.Server.MapPath([String].Format("~/{0}/{1}/{2}", LogFileDir, datePath, projektNazwa))

                If Not Directory.Exists(fileLogDirectory) Then
                    Directory.CreateDirectory(fileLogDirectory)
                End If

                Try
                    Using fs As New FileStream(fileLogDirectory, FileMode.Create)
                        Using bw As New BinaryWriter(fs)
                            bw.Write(xmlFileText)
                            bw.Close()
                        End Using
                        retVal = fs.Length
                    End Using
                Catch ex As Exception
                    Dim msg = ex.Message
                    msg = "Problem z dostępem do pliku/foleru. (" & ex.Message.Substring(0, If((msg.Length < 200), msg.Length, 200))
                    Return msg
                End Try

            End If

            Return retVal

        End Function
#End Region

#Region "Funkcje pomocnicze"
        Public Shared Function IsTextAfterTrim(text1 As String, ParamArray texts As String()) As Boolean
            If IsTextAfterTrim(text1) Then
                For Each s As String In texts
                    If Not IsTextAfterTrim(s) Then
                        Return False
                    End If
                Next
                Return True
            End If
            Return False
        End Function
#End Region

    End Class

#End Region
End Namespace