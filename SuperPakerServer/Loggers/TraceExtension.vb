Imports System
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.IO


Public Class LogDirectory
    Private CONST_SCIEZKA_LOG As String = "C:\\Projekty\\SuperPaker\\Pliki\\Trace\" 'Zmieniać tylko tą ścieżkę


    Public ReadOnly Property LogFilePath() As String
        Get
            Dim now As DateTime = DateTime.Now
            Return CONST_SCIEZKA_LOG & "\" & now.Year & "\" & now.Month & "\"
        End Get
    End Property

    Public Sub LogDirectory()
        Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath))
    End Sub

    Public CONST_WRITE_IN As Boolean = True
    Public CONST_WRITE_OUT As Boolean = True
    Private PLIK_LOGOWANIA_WS_EXT As String = ".log"
    Private PLIK_LOGOWANIA_LOG As String = "Log.log"

    Public Function GetOdpowiedz(ByVal wsName As String)
        LogDirectory()

        Return Path.Combine(LogFilePath, Now.Day & "_" & Now.Hour & "_" & wsName & PLIK_LOGOWANIA_WS_EXT)
    End Function
    Public Function GetLogFilename()
        LogDirectory()
        Return Path.Combine(LogFilePath, Now.Day & "_" & Now.Hour & "_" & PLIK_LOGOWANIA_LOG)
    End Function
End Class

' Define a SOAP Extension that traces the SOAP request and SOAP response
' for the XML Web service method the SOAP extension is applied to.
Public Class TraceExtension
    Inherits SoapExtension

    Private oldStream As Stream
    Private newStream As Stream
    Private m_filename As String



    ' Save the Stream representing the SOAP request or SOAP response into
    ' a local memory buffer.
    Public Overrides Function ChainStream(ByVal stream As Stream) As Stream
        oldStream = stream
        newStream = New MemoryStream()
        Return newStream
    End Function

    ' When the SOAP extension is accessed for the first time, the XML Web
    ' service method it is applied to is accessed to store the file
    ' name passed in, using the corresponding SoapExtensionAttribute.    
    Public Overloads Overrides Function GetInitializer(ByVal methodInfo As  _
        LogicalMethodInfo, _
    ByVal attribute As SoapExtensionAttribute) As Object
        Return CType(attribute, TraceExtensionAttribute).Filename
    End Function

    ' The SOAP extension was configured to run using a configuration file
    ' instead of an attribute applied to a specific XML Web service
    ' method.  Return a file name based on the class implementing the Web
    ' Service's type.
    Public Overloads Overrides Function GetInitializer(ByVal WebServiceType As  _
      Type) As Object
        ' Return a file name to log the trace information to, based on the
        ' type.
        Dim ustawienia As New LogDirectory
        Return ustawienia.GetOdpowiedz(WebServiceType.FullName)
    End Function

    ' Receive the file name stored by GetInitializer and store it in a
    ' member variable for this specific instance.
    Public Overrides Sub Initialize(ByVal initializer As Object)
        m_filename = CStr(initializer)
    End Sub

    ' If the SoapMessageStage is such that the SoapRequest or SoapResponse
    ' is still in the SOAP format to be sent or received over the network,
    ' save it out to file.
    Public Overrides Sub ProcessMessage(ByVal message As SoapMessage)
        Select Case message.Stage
            Case SoapMessageStage.BeforeSerialize
            Case SoapMessageStage.AfterSerialize
                WriteOutput(message)
            Case SoapMessageStage.BeforeDeserialize
                WriteInput(message)
            Case SoapMessageStage.AfterDeserialize
        End Select
    End Sub

    ' Write the SOAP message out to a file.
    Public Sub WriteOutput(ByVal message As SoapMessage)
        newStream.Position = 0
        Dim ustawienia As New LogDirectory

        If ustawienia.CONST_WRITE_OUT = True Then
            Dim fs As New FileStream(m_filename, FileMode.Append, _
                                     FileAccess.Write)
            Dim w As New StreamWriter(fs)
            w.WriteLine("-----Response at " + DateTime.Now.ToString())
            w.Flush()
            Copy(newStream, fs)
            w.Close()
            newStream.Position = 0
        End If

        Copy(newStream, oldStream)
    End Sub

    ' Write the SOAP message out to a file.
    Public Sub WriteInput(ByVal message As SoapMessage)
        Copy(oldStream, newStream)
        Dim ustawienia As New LogDirectory

        If ustawienia.CONST_WRITE_IN = True Then
            Dim fs As New FileStream(m_filename, FileMode.Append, _
                                     FileAccess.Write)
            Dim w As New StreamWriter(fs)

            Dim ADRES As String = ""

            If Not HttpContext.Current.Request.UserHostAddress Is Nothing Then
                ADRES = HttpContext.Current.Request.UserHostAddress
            End If

            w.WriteLine("----- Request at " + DateTime.Now.ToString() + " Address: " + ADRES)
            w.Flush()
            newStream.Position = 0
            Copy(newStream, fs)
            w.Close()
        End If
        newStream.Position = 0
    End Sub

    Sub Copy(ByVal fromStream As Stream, ByVal toStream As Stream)
        Dim reader As New StreamReader(fromStream)
        Dim writer As New StreamWriter(toStream)
        writer.WriteLine(reader.ReadToEnd())
        writer.Flush()
    End Sub
End Class

' Create a SoapExtensionAttribute for our SOAP Extension that can be
' applied to an XML Web service method.
<AttributeUsage(AttributeTargets.Method)> _
Public Class TraceExtensionAttribute
    Inherits SoapExtensionAttribute

    Dim ustawienia As New LogDirectory
    Private m_filename As String = ustawienia.GetLogFilename()
    Private m_priority As Integer

    Public Overrides ReadOnly Property ExtensionType() As Type
        Get
            Return GetType(TraceExtension)
        End Get
    End Property

    Public Overrides Property Priority() As Integer
        Get
            Return m_priority
        End Get
        Set(ByVal Value As Integer)
            m_priority = value
        End Set
    End Property

    Public Property Filename() As String
        Get
            Return m_filename
        End Get
        Set(ByVal Value As String)
            m_filename = value
        End Set
    End Property
End Class
