#Region "Imports"
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Xml
Imports System.Xml.Schema
Imports System.IO
Imports System.Data.SqlClient
Imports SuperPakerServer.DataAccess
#End Region

<System.Web.Services.WebService(Namespace:="http://superpaker-test.cursor.pl/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class SuperPakerService
    Inherits System.Web.Services.WebService
    Private Url As String
    'Private Shared ReadOnly logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private Const CONST_SCIEZKA_PLIKOW As String = "C:\\Projekty\\SuperPaker\\ZalacznikiAplikacji\\ProjektId_{0}\\"
    Private Const CONST_SCIEZKA_PLIKOW_SZABLONOW As String = "C:\\Projekty\\SuperPaker\\ZalacznikiAplikacji\\Szablony\\"

    Public kontaktIt = " Przekaż ten problem do systemu zgłoszeń działu IT firmy Cursor: it@cursor.pl"
    'Public connectionString = "Data Source=NBG-SRV-31;Initial Catalog=TEST_SUPERPAKER;User Id=TEST_SUPER_PAKER_USER;Password=ofo7s4u56us757jcu456us4;Connection Timeout=240;"
    Public connectionString = "Data Source=10.1.70.31;Initial Catalog=TEST_SUPERPAKER;User Id=TEST_SUPER_PAKER_USER;Password=ofo7s4u56us757jcu456us4;Connection Timeout=240;"

    Private msg As String = ""

    Public Sub New()
        DataAccess.Config.ConnectionStringSuperPaker = connectionString
        'DPDWSDao.Config.ConnectionStringQGUAR = connectionStringQguar
    End Sub

#Region "Logowanie"

    ''' <summary>
    ''' Logowanie użytkownika aplikacji do systemu 
    ''' </summary>
    ''' <param name="login">login użytkownika</param>
    ''' <param name="haslo">hasło</param>
    ''' <param name="wersja">wersja aplikacji</param>
    ''' <param name="komputer">komputer z którego wykonywane jest logowanie</param>
    ''' <returns>Dane zalogowanego użytkownika i jego uprawnienia </returns>
    ''' <remarks></remarks>
    ''' 
    <WebMethod()> _
    Public Function Zaloguj(ByVal login As String, ByVal haslo As String, ByVal wersja As String, _
           ByVal komputer As String) As ZalogujWynik
        Dim wynik As New ZalogujWynik
        Dim cnn As SqlConnection

        'łączymy do bazy danych
        Try
            cnn = New SqlConnection()
            cnn.ConnectionString = ConnectionStringSuperPaker
            cnn.Open()
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd połączenia do bazy danych: " & ex.Message & vbNewLine & kontaktIt
            Return wynik
        End Try

        'wywołujemy procedurę zaloguj
        Dim cmd As New SqlClient.SqlCommand("UP_UZYTKOWNIK_ZALOGUJ", cnn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 6
        cmd.Parameters.AddWithValue("@login", login)
        cmd.Parameters.AddWithValue("@haslo", haslo)
        cmd.Parameters.AddWithValue("@wersja", wersja)
        cmd.Parameters.AddWithValue("@komputer", komputer)
        cmd.Parameters.AddWithValue("@adres", Me.Context.Request.UserHostAddress)
        cmd.Parameters.Add("@sesja", SqlDbType.VarBinary, 16).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@uzytkownik_id", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@uzytkownik", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@telefon", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@rola_id", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@rola", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@status", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@status_opis", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@czy_pierwszy", SqlDbType.Int).Direction = ParameterDirection.Output

        cmd.Parameters.Add("@CULTURE_CODE", SqlDbType.NVarChar, 5).Direction = ParameterDirection.Output
        Try
            cmd.ExecuteNonQuery()

            Select Case cnn.Database
                Case "TEST_SUPERPAKER"
                    wynik.adresPolaczenia = "Środowisko: Testowe"
                Case "SUPERPAKER_PROD"
                    wynik.adresPolaczenia = "Środowisko: Produkcyjne"
                Case Else
                    wynik.adresPolaczenia = "Środowisko: Nieznane!"
            End Select

        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd komunikacji z bazą: " & ex.Message & kontaktIt
            cnn.Close()
            Return wynik
        End Try

        wynik.status = cmd.Parameters("@status").Value
        wynik.status_opis = cmd.Parameters("@status_opis").Value
        If wynik.status <> -1 Then
            wynik.sesja = cmd.Parameters("@sesja").Value
            wynik.uzytkownik_id = cmd.Parameters("@uzytkownik_id").Value
            wynik.uzytkownik = IIf(IsDBNull(cmd.Parameters("@uzytkownik").Value), "", cmd.Parameters("@uzytkownik").Value)
            wynik.telefon = IIf(IsDBNull(cmd.Parameters("@telefon").Value), "", cmd.Parameters("@telefon").Value)
            wynik.czy_pierwszy = IIf(IsDBNull(cmd.Parameters("@czy_pierwszy").Value), True, cmd.Parameters("@czy_pierwszy").Value)
            wynik.kodJezyka = IIf(IsDBNull(cmd.Parameters("@CULTURE_CODE").Value), "pl-PL", cmd.Parameters("@CULTURE_CODE").Value)
        End If
        cnn.Close()
        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZmienHaslo(ByVal sesja As Byte(), ByVal obecne_haslo As String, ByVal nowe_haslo As String) As ZmienHasloWynik
        Dim wynik As New ZmienHasloWynik
        Dim cnn As SqlConnection

        'łączymy do bazy
        Try
            cnn = New SqlConnection()
            cnn.ConnectionString = ConnectionStringSuperPaker
            cnn.Open()
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd połączenia do bazy danych: " & ex.Message & vbNewLine & kontaktIt
            '   logger.Error("ZmienHasloUzytkownika:Błąd połączenia do bazy danych: ", ex)
            Return wynik
        End Try

        'wywołujemy procedurę
        Dim cmd As New SqlClient.SqlCommand("UP_UZYTKOWNIK_ZMIEN_HASLO", cnn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@sesja", sesja)
        cmd.Parameters.AddWithValue("@OBECNE_HASLO_IN", obecne_haslo)
        cmd.Parameters.AddWithValue("@NOWE_HASLO_IN", nowe_haslo)
        cmd.Parameters.Add("@status", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@status_opis", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd komunikacji z bazą: " & ex.Message & kontaktIt
            cnn.Close()
            '  logger.Error("ZmienHasloUzytkownika:Błąd komunikacji z bazą: ", ex)
            Return wynik
        End Try

        wynik.status = cmd.Parameters("@status").Value
        wynik.status_opis = cmd.Parameters("@status_opis").Value
        cnn.Close()
        Return wynik
    End Function

#End Region

#Region "Grupy projektowe"

    <WebMethod()> _
    Public Function UzytkownikGrupaProjektowaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As UzytkownikGrupaProjektowaListaPobierzWynik

        Dim wynik As New UzytkownikGrupaProjektowaListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_GRUPA_PROJEKTOW_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function UzytkownikGrupaProjektowaDanePobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal grupaProjektowaId As Integer) As UzytkownikGrupaProjektowaDanePobierzWynik

        Dim wynik As New UzytkownikGrupaProjektowaDanePobierzWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@GRUPA_ID", grupaProjektowaId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_UZYTKOWNIK_GRUPA_PROJEKTOW_DANE_GRUPY_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikGrupaProjektowaZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal grupaProjektowa As UzytkownikGrupaProjektowa) As UzytkownikGrupaProjektowaZapiszWynik

        Dim wynik As New UzytkownikGrupaProjektowaZapiszWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim grupaIdOut As Integer

        If IsNothing(grupaProjektowa) = True Then
            wynik.status = -1
            wynik.status_opis = "Nie podano danych do przetworzenia"
            Return wynik
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim lp As Integer = 1
            Dim projektyDt As DataTable = getListaTypeTbl()
            If IsNothing(grupaProjektowa.ProjektListaId) = False Then
                For Each i As Integer In grupaProjektowa.ProjektListaId
                    Dim pR As DataRow = projektyDt.NewRow
                    pR([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.ID)) = lp
                    pR([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_INT)) = i
                    lp += 1
                    projektyDt.Rows.Add(pR)
                Next
                projektyDt.AcceptChanges()
            End If

            Dim uzytkownicyDt As DataTable = getListaTypeTbl()
            lp = 1
            If IsNothing(grupaProjektowa.UzytkownikListaId) = False Then
                For Each i As Integer In grupaProjektowa.UzytkownikListaId
                    Dim uR As DataRow = uzytkownicyDt.NewRow
                    uR([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.ID)) = lp
                    uR([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_INT)) = i
                    lp += 1
                    uzytkownicyDt.Rows.Add(uR)
                Next
                uzytkownicyDt.AcceptChanges()
            End If

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@PROJEKTY_LISTA_IN", projektyDt)
            htParams.Add("@UZYTKOWNICY_LISTA_IN", uzytkownicyDt)
            htParams.Add("@NAZWA_GRUPY", grupaProjektowa.nazwa)
            htParams.Add("@OPIS", grupaProjektowa.opis)
            htParams.Add("@GRUPA_ID", grupaProjektowa.grupaId)
            htParams.Add("@OUT_GRUPA_ID", grupaIdOut)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_UZYTKOWNIK_GRUPA_PROJEKTOW_EDYTUJ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If procRes.Status = 0 Then
                    wynik.grupaId = NZ(htParams("@OUT_GRUPA_ID"), 0)
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Uzytkownik"

    <WebMethod()> _
    Public Function UzytkownikListaPobierz(ByVal sesja As Byte()) As UzytkownikListaWynik

        Dim wynik As New UzytkownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikDanePobierz(ByVal sesja As Byte(), ByVal uzytkownikId As Integer) As UzytkownikInfoWynik

        Dim wynik As New UzytkownikInfoWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownikId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_DANE_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikEdytuj(ByVal sesja As Byte(), ByVal uzytkownikEdytowanyId As Integer) As UzytkownikEdytujWynik

        Dim wynik As New UzytkownikEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_EDYTOWANY_IN", uzytkownikEdytowanyId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_UZYTKOWNIK_EDYCJA_ROZPOCZNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikEdytujZapisz(ByVal sesja As Byte(), ByVal dane As DataSet, ByVal strHaslo1 As String, ByVal strHaslo2 As String, ByVal bitHasloZmiana As Boolean, ByVal cultureCode As String) As UzytkownikZapiszWynik

        Dim wynik As New UzytkownikZapiszWynik

        If dane.Tables.Count > 1 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim uzytkownikTable As New DataTable '
            Dim roleTable As New DataTable
            Dim projektyTable As New DataTable
            Dim funkcjeTable As New DataTable
            Dim uzytkownikIdEtytowanyOUT As Integer = 0

            If dane.Tables.Count > 0 Then
                uzytkownikTable = dane.Tables(0).Copy
            End If

            If dane.Tables.Count > 1 Then
                roleTable = dane.Tables(1).Copy
            End If

            If dane.Tables.Count > 2 Then
                projektyTable = dane.Tables(2).Copy
            End If

            If dane.Tables.Count > 3 Then
                funkcjeTable = dane.Tables(3).Copy
            End If

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@HASLO1_IN", strHaslo1)
                htParams.Add("@HASLO2_IN", strHaslo2)
                htParams.Add("@HASLO_ZMIANA_IN", bitHasloZmiana)
                htParams.Add("@CULTURE_CODE_ID_IN", cultureCode)
                htParams.Add("@TableUzytkownik_in", uzytkownikTable)
                htParams.Add("@TableRole_in", roleTable)
                htParams.Add("@TableProjekty_in", projektyTable)
                htParams.Add("@TableFunkcje_in", funkcjeTable)
                htParams.Add("@OUT_UZYTKOWNIK_ID_EDYTOWANY", uzytkownikIdEtytowanyOUT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_UZYTKOWNIK_EDYCJA_ZAPISZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.uzytkownikId = NZ(htParams("@OUT_UZYTKOWNIK_ID_EDYTOWANY"), 0)
                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function UzytkownikEdytujAnuluj(ByVal sesja As Byte(), ByVal uzytkownik_id As Integer)

        Dim wynik As New BlokadaUsunWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ID_BLOKOWANE", uzytkownik_id)
            htParams.Add("@TABELA", "UZYTKOWNIK")

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_TABELA_ODBLOKUJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikUstawDateWaznosci(ByVal sesja As Byte(), ByVal uzytkownikId As Integer, ByVal dataWaznosci As DateTime) As StatusWynik

        Dim wynik As New StatusWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownikId)
            htParams.Add("@DATA_DO_IN", dataWaznosci)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_UZYTKOWNIK_USTAW_DATE_WAZNOSCI", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik

    End Function

    <WebMethod()> _
    Public Function UserDSAtrybutListaPobierz(ByVal sesja As Byte()) As SlownikItemsWynik

        Dim wynik As New SlownikItemsWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_DS_ATRYBUTY_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikNotyfikacjeEmailDanePobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal uzytkownikId As Integer) As UzytkownikNotyfikacjeEmailDanePobierzWynik

        Dim wynik As New UzytkownikNotyfikacjeEmailDanePobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownikId)
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_NOTYFIKACJE_EMAIL", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function UzytkownikNotyfikacjeEmailDaneZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As UzytkownikNotyfikacjeEmailDaneZapiszWynik

        Dim wynik As New UzytkownikNotyfikacjeEmailDaneZapiszWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim notyfikacjeTable As New DataTable '

            If dane.Tables.Count > 0 Then
                notyfikacjeTable = dane.Tables(0).Copy
            End If


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@TableNotyfikacje_in", notyfikacjeTable)
                htParams.Add("@PROJEKT_ID_IN", projektId)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_UZYTKOWNIK_NOTYFIKACJE_EMAIL_ZMIEN", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                  
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function JezykiListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_CULTURE_CODE_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Magazyny"

    <WebMethod()> _
    Public Function MagazynyListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As MagazynyListaWynik

        Dim wynik As New MagazynyListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim mCount As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@OUT_MAGAZYNY_LICZNIK", mCount)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_MAGAZYN_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If wynik.status = 0 Then
                    wynik.magazyny_licznik = mCount
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Magazyny wirtualne"
    <WebMethod()> _
    Public Function MagazynWirtualnyListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal aktywny As Boolean) As MagazynWirtualnyListaWynik

        Dim wynik As New MagazynWirtualnyListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@AKTYWNY", aktywny)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_MAGAZYN_WIRTUALNY_LISTA", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function MagazynWirtualnyDodaj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal nazwa As String, ByVal opis As String) As MagazynWirtualnyDodajWynik

        Dim wynik As New MagazynWirtualnyDodajWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""
            Dim magazynIdOUT As Integer = 0

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@MAG_WIRT_NAZWA", nazwa)
            htParams.Add("@MAG_WIRT_OPIS", opis)
            htParams.Add("@MAGAZYN_WIRTUALNY_ID", magazynIdOUT)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_MAGAZYN_WIRTUALNY_DODAJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If wynik.status = 0 Then
                    wynik.magazyn_wirtualny_id = NZ(htParams("@MAGAZYN_WIRTUALNY_ID"), "")
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function MagazynWirtualnyEdytuj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal magazyn_nazwa_stara As String, ByVal magazyn_nazwa_nowa As String) As MagazynWirtualnyEdytujWynik

        Dim wynik As New MagazynWirtualnyEdytujWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@MAG_WIRT_NAZWA_STARA", magazyn_nazwa_stara)
            htParams.Add("@MAG_WIRT_NAZWA_NOWA", magazyn_nazwa_nowa)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_MAGAZYN_WIRTUALNY_EDYTUJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function MagazynWirtualnyMMDodaj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal magazyn_nazwa_zrodlo As String, ByVal magazyn_nazwa_cel As String, ByVal dane As DataSet, ByVal info As String, ByVal dok_nr As String) As MagazynWirtualnyMMDodajWynik

        Dim wynik As New MagazynWirtualnyMMDodajWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim ProduktTable As DataTable

            If dane.Tables.Count > 0 Then
                ProduktTable = dane.Tables(0).Copy
            Else
                ProduktTable = New DataTable()
            End If

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PRODUKT_LISTA", ProduktTable)
                htParams.Add("@PROJEKT_ID", projekt_id)
                htParams.Add("@MAG_WIRT_NAZWA_ZRODLO", magazyn_nazwa_zrodlo)
                htParams.Add("@MAG_WIRT_NAZWA_CEL", magazyn_nazwa_cel)
                htParams.Add("@INFO", info)
                htParams.Add("@DOK_NR", dok_nr)
                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_MAGAZYN_WIRTUALNY_MM_DODAJ", htParams)
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If



        Return wynik

    End Function
#End Region

#Region "Projekty"

    <WebMethod()> _
    Public Function ProjektyListaPobierz(ByVal uzytkownik_id As Integer, ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownik_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PROJEKT_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProjektyListaAtrPobierz(ByVal uzytkownik_id As Integer, ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownik_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PROJEKT_LISTA_ATR_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProjektParametrListaPobierz(ByVal sesja As Byte(), ByVal projektID As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektID)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PROJEKT_PARAMETR_LISTA", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProjektEnovaDanePobierz(ByVal sesja As Byte(), ByVal projektID As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PROJEKT_ENOVA_DANE_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


#End Region

#Region "Role"

    <WebMethod()> _
    Public Function RoleListaPobierz(ByVal uzytkownik_id As Integer, ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownik_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ROLA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Zamowienia"

    <WebMethod()> _
    Public Function ZamowienieListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal statusyDS As DataSet) As ZamowienieListaWynik
        'Piotr - dodana mozliwosc filtrowania po statusach
        Dim wynik As New ZamowienieListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable

        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    ' htParams.Add("@ZAMOWIENIE_STATUS_ID_IN", zamowienieStatusId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_LISTA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra"
        End If

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer,
                                      ByVal sortPo As String,
                                      ByVal sortAsc As Boolean,
                                      ByVal statusyDS As DataSet) As ZamowienieStronaWynik

        Dim wynik As New ZamowienieStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0
        Dim dTableStatusy As DataTable = New DataTable

        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    'htParams.Add("@ZAMOWIENIE_STATUS_ID_IN", zamowienieStatusId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STRONA_NR_IN", stronaNumer)
                    htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                    htParams.Add("@SORT_PO_IN", sortPo)
                    htParams.Add("@SORT_KIERUNEK_IN", sortAsc)

                    htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_ZAMOWIENIE_STRONA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    If procRes.Status = 0 Then
                        wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                    End If
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieListaFiltrListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_LISTA_FILTR_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()>
    Public Function ZamowieniaListaPobierzOdDaty(ByVal sesja As Byte(), projektId As Integer, dateOd As Date) As ZamowienieDaneWynik

        Dim wynik As New ZamowienieDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DATE_OD", dateOd)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIA_LISTA_POBIERZ_OD_DATY", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()>
    Public Function EdytujPozycjeZamowienia(ByVal sesja As Byte(), projektId As Integer, ByVal dane As DataSet) As ZamowienieEdycjaStatusWynik

        Dim wynik As New ZamowienieEdycjaStatusWynik
        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then
            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

                Dim outStr As String = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@TableEdycja_in", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_EDYCJA_POZYCJI_LISTA_XML", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else

            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane zamówień do sprawdzenia."
        End If

        Return wynik


    End Function


    <WebMethod()> _
    Public Function ZamowienieKompletacjaDanePobierz(ByVal sesja As Byte(), ByVal zamowienieId As Integer) As ZamowienieDaneWynik

        Dim wynik As New ZamowienieDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieDanePobierz(ByVal sesja As Byte(), ByVal zamowienieId As Integer) As ZamowienieDaneWynik

        Dim wynik As New ZamowienieDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieEdytuj(ByVal sesja As Byte(), ByVal zamowienieEdytowaneId As Integer) As ZamowienieEdytujWynik

        Dim wynik As New ZamowienieEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_EDYTOWANE_IN", zamowienieEdytowaneId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_EDYCJA_ROZPOCZNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function AdresEdycjaZapisz(ByVal sesja As Byte(), ByVal adres_id As Integer, ByVal imie As String, _
                                           ByVal nazwisko As String, ByVal nazwa As String, ByVal firma As String, ByVal ulica As String, _
                                           ByVal nr_domu As String, ByVal nr_mieszkania As String, ByVal urzad_pocztowy As String, ByVal kraj_region As String, _
                                           ByVal wojewodztwo As String, ByVal kod_pocztowy As String, ByVal miasto As String, ByVal kraj As String, _
                                           ByVal e_mail As String, ByVal telefon As String, ByVal telefon_komorkowy As String, ByVal fax As String, _
                                           ByVal nip As String, ByVal numer_vat As String, ByVal data_od As DateTime, ByVal data_do As DateTime, _
                                           ByVal zrodlo_zmian As String, ByVal opis_zmian As String) As AdresEdycjaZapiszWynik

        Dim wynik As New AdresEdycjaZapiszWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ADRES_ID", NZ(adres_id, DBNull.Value))
            htParams.Add("@IMIE", NZ(imie, DBNull.Value))
            htParams.Add("@NAZWISKO", NZ(nazwisko, DBNull.Value))
            htParams.Add("@NAZWA", NZ(nazwa, DBNull.Value))
            htParams.Add("@FIRMA", NZ(firma, DBNull.Value))
            htParams.Add("@ULICA", NZ(ulica, DBNull.Value))
            htParams.Add("@NR_DOMU", NZ(nr_domu, DBNull.Value))
            htParams.Add("@NR_MIESZKANIA", NZ(nr_mieszkania, DBNull.Value))
            htParams.Add("@URZAD_POCZTOWY", NZ(urzad_pocztowy, DBNull.Value))
            htParams.Add("@KRAJ_REGION", NZ(kraj_region, DBNull.Value))
            htParams.Add("@WOJEWODZTWO", NZ(wojewodztwo, DBNull.Value))
            htParams.Add("@KOD_POCZTOWY", NZ(kod_pocztowy, DBNull.Value))
            htParams.Add("@MIASTO", NZ(miasto, DBNull.Value))
            htParams.Add("@KRAJ", NZ(kraj, DBNull.Value))
            htParams.Add("@E_MAIL", NZ(e_mail, DBNull.Value))
            htParams.Add("@TELEFON", NZ(telefon, DBNull.Value))
            htParams.Add("@TELEFON_KOMORKOWY", NZ(telefon_komorkowy, DBNull.Value))
            htParams.Add("@FAX", NZ(fax, DBNull.Value))
            htParams.Add("@NIP", NZ(nip, DBNull.Value))
            htParams.Add("@NUMER_VAT", NZ(numer_vat, DBNull.Value))
            htParams.Add("@DATA_OD", IIf(data_od = Nothing, DBNull.Value, data_od))
            htParams.Add("@DATA_DO", IIf(data_do = Nothing, DBNull.Value, data_do))
            htParams.Add("@ZRODLO_ZMIAN", NZ(zrodlo_zmian, DBNull.Value))
            htParams.Add("@OPIS_ZMIAN", NZ(opis_zmian, DBNull.Value))


            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ADRES_EDYCJA_ZAPISZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieEdytujAnuluj(ByVal sesja As Byte(), ByVal zamowienie_id As Integer)

        Dim wynik As New BlokadaUsunWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ID_BLOKOWANE", zamowienie_id)
            htParams.Add("@TABELA", "UZYTKOWNIK")

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_TABELA_ODBLOKUJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik

    End Function

    ''' <summary>
    ''' Metoda webserwisu walidująca i zapisująca zamówienia w bazie
    ''' </summary>
    ''' <param name="sesja">identyfikator sesji</param>
    ''' <param name="klientId">identyfikator klienta</param>
    ''' <param name="zamowienieTypId">identyfikator źródła pochodzenia zamówienia (wewnętrzny)</param>
    ''' <param name="dane">DataSet z zamówieniami i pozycjami z poszczególnych zamówień</param>
    ''' <returns>Informacja czy zamówienia przeszły walidację merytoryczną i zostały poprawnie zapisane</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function zlozZamowienie(ByVal sesja As Byte(), ByVal klientId As Integer, ByVal projektId As Integer, ByVal zamowienieTypId As Integer, ByVal dane As DataSet) As ZamowienieZlozWynik
        Dim wynik As New ZamowienieZlozWynik
        Dim sData As New DataSet 'dataset ze statusem walidacji merytorycznej per zamówienie

        If dane.Tables.Count > 1 AndAlso dane.Tables(0).Rows.Count > 0 AndAlso dane.Tables(1).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable
            Dim zamowieniePozycjaTable As New DataTable
            Dim zamowienieAtrybutTable As DataTable

            If dane.Tables.Count > 0 Then
                zamowienieTable = dane.Tables(0).Copy
            End If

            If dane.Tables.Count > 1 Then
                zamowieniePozycjaTable = dane.Tables(1).Copy
            End If


            If dane.Tables.Count > 2 Then
                zamowienieAtrybutTable = dane.Tables(2).Copy
            Else
                zamowienieAtrybutTable = New DataTable()
            End If

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@KLIENT_ID_IN", klientId)
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_TYP_ID", zamowienieTypId)
                htParams.Add("@TableZamowienie_in", zamowienieTable)
                htParams.Add("@TableZamowieniePozycja_in", zamowieniePozycjaTable)
                If dane.Tables.Count > 2 Then
                    htParams.Add("@TableZamowienieAtrybut_in", zamowienieAtrybutTable)
                End If

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "STG_ZAMOWIENIE_STG_ZAPISZ", htParams, sData)
                    wynik.dane = sData
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = sData
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieStatusListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As ZamowienieStatusListaWynik
        Dim wynik As New ZamowienieStatusListaWynik
        Dim sData As New DataSet

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@TableZamowienie_in", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_STATUS_LISTA", htParams, sData)
                    wynik.dane = sData
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieStatusKurierListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As ZamowienieStatusListaWynik
        Dim wynik As New ZamowienieStatusListaWynik
        Dim sData As New DataSet

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@TableZamowienie_in", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_STATUS_KURIER_LISTA", htParams, sData)
                    wynik.dane = sData
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane zamówień do sprawdzenia."
        End If


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieStatusKurierListaXmlPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As ZamowienieStatusListaXmlWynik
        Dim wynik As New ZamowienieStatusListaXmlWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

                Dim outStr As String = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@TableZamowienie_in", zamowienieTable)
                htParams.Add("@OUT_ZAMOWIENIE_STATUSY_XML:-1", outStr)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_STATUS_KURIER_LISTA_XML", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.xml = NZ(htParams("@OUT_ZAMOWIENIE_STATUSY_XML:-1"), "")

                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane zamówień do sprawdzenia."
        End If


        Return wynik
    End Function

    <WebMethod()> _
    Public Function StanMagazynowyPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As StanMagazynowyPobierzWynik
        Dim wynik As New StanMagazynowyPobierzWynik
        Dim sData As New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_STANY_Q_SLOWNIK_PRODUKTOW", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function StanMagazynowyUszkodzonePobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As StanMagazynowyUszkodzonePobierzWynik
        Dim wynik As New StanMagazynowyUszkodzonePobierzWynik
        Dim sData As New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_STAN_PRODUKTOW_USZKODZONYCH", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


    <WebMethod()> _
    Public Function StanMagazynowyMagazynWirtualnyPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal magazynyWirtualne As DataSet, ByVal czyWirtualny As Boolean, ByVal czyStatusJakosci As Boolean) As StanMagazynowyMagazynWirtualnyPobierzWynik
        Dim wynik As New StanMagazynowyMagazynWirtualnyPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableMagazyny As DataTable = New DataTable

        If magazynyWirtualne.Tables.Count > 0 Then
            dTableMagazyny = magazynyWirtualne.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@MAGAZYN_WIRTUALNY_LISTA", dTableMagazyny)
                htParams.Add("@CZY_WIRTUALNY", czyWirtualny)
                htParams.Add("@CZY_STATUS_JAKOSCI", czyStatusJakosci)

                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_STANY_Q_MAGAZYN_WIRTUALNY", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik

    End Function

    <WebMethod()> _
    Public Function StanMagazynowyMagazynWirtualnyProduktListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal produkty As DataSet, ByVal magazynyWirtualne As DataSet, ByVal czyWirtualny As Boolean, ByVal czyStatusJakosci As Boolean) As StanMagazynowyMagazynWirtualnyPobierzWynik
        Dim wynik As New StanMagazynowyMagazynWirtualnyPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableMagazyny As DataTable = New DataTable
        Dim dTableProdukty As DataTable = New DataTable

        If produkty.Tables.Count > 0 Then
            dTableProdukty = produkty.Tables(0).Copy
        End If

        If magazynyWirtualne.Tables.Count > 0 Then
            dTableMagazyny = magazynyWirtualne.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@PRODUKT_LISTA", dTableProdukty)
                htParams.Add("@MAGAZYN_WIRTUALNY_LISTA", dTableMagazyny)
                htParams.Add("@CZY_WIRTUALNY", czyWirtualny)
                htParams.Add("@CZY_STATUS_JAKOSCI", czyStatusJakosci)

                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_STANY_Q_MAGAZYN_WIRTUALNY_PRODUKT_LISTA", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik

    End Function

    <WebMethod()> _
    Public Function StanMagazynowyProduktListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal produkty As DataSet, ByVal magazynyWirtualne As DataSet, ByVal czyWirtualny As Boolean, ByVal czyStatusJakosci As Boolean) As StanMagazynowyProduktListaPobierzWynik
        Dim wynik As New StanMagazynowyProduktListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableMagazyny As DataTable = New DataTable
        Dim dTableProdukty As DataTable = New DataTable

        If produkty.Tables.Count > 0 Then
            dTableProdukty = produkty.Tables(0).Copy
        End If

        If magazynyWirtualne.Tables.Count > 0 Then
            dTableMagazyny = magazynyWirtualne.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@PRODUKT_LISTA", dTableProdukty)
                htParams.Add("@MAGAZYN_WIRTUALNY_LISTA", dTableMagazyny)
                htParams.Add("@CZY_WIRTUALNY", czyWirtualny)
                htParams.Add("@CZY_STATUS_JAKOSCI", czyStatusJakosci)

                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_STANY_Q_MAGAZYN_WIRTUALNY_STANY_BIEZACE_PRODUKT_LISTA", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik

    End Function


    <WebMethod()> _
    Public Function AdresEdycjaWarunekSprawdz(ByVal sesja As Byte(), ByVal adres_id As Integer) As AdresEdycjaWarunekSprawdzWynik
        Dim wynik As New AdresEdycjaWarunekSprawdzWynik
        Dim sData As New DataSet




        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""
            Dim OutMoznaEdytowac As Integer = 0
            Dim OutMoznaOpis As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ADRES_ID", adres_id)
            htParams.Add("@OUT_MOZNA_EDYTOWAC", OutMoznaEdytowac)
            htParams.Add("@OUT_MOZNA_EDYTOWAC_OPIS:255", OutMoznaOpis)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ADRES_EDYCJA_WARUNEK_SPRAWDZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.OutMoznaEdytowac = NZ(htParams("@OUT_MOZNA_EDYTOWAC"), -1)
                    wynik.OutMoznaEdytowacOpis = NZ(htParams("@OUT_MOZNA_EDYTOWAC_OPIS:255"), "")

                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using



        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieStatusyListaPobierz(ByVal sesja As Byte(), projektId As Integer, filtr As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@FILTR_STATUSOW_IN", filtr)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_STATUSY_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function PlatnoscMetodaListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PLATNOSC_METODA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieNrAnulujPojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieNr As String) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", zamowienieNr)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_NR_ANULUJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieAnulujPojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_ANULUJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieAnulujLista(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As StatusWynik

        Dim wynik As New StatusWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_LISTA_IN", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_LISTA_ANULUJ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowa lista zamówień do anulowania"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieUsunPojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_USUN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowieniePozycjaUsunPojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal pozycjaId As Integer, ByVal komentarz As String, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@KOMENTARZ_IN", komentarz)
            htParams.Add("@ZAMOWIENIE_POZYCJA_ID_IN", pozycjaId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_POZYCJE_USUN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
  

    <WebMethod()> _
    Public Function ZamowienieOznaczSpersonalizowanePojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_PERSONALIZACJA_ZATWIERDZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieOznaczSpersonalizowaneLista(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As StatusWynik

        Dim wynik As New StatusWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_LISTA_IN", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_PERSONALIZACJA_LISTA_ZATWIERDZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowa lista zamówień do oznaczenia jako spersonalizowane"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieOznaczOplaconePojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_PRZELEW_ZATWIERDZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieOznaczOplaconeLista(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As StatusWynik

        Dim wynik As New StatusWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_LISTA_IN", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_PRZELEW_LISTA_ZATWIERDZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowa lista zamówień do oznaczenia jako spersonalizowane"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieOznaczDostarczonoPojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_DORECZENIE_ZATWIERDZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieOznaczDostarczonoLista(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As StatusWynik

        Dim wynik As New StatusWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_LISTA_IN", zamowienieTable)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_DORECZENIE_LISTA_ZATWIERDZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowa lista zamówień do oznaczenia jako spersonalizowane"
        End If

        Return wynik
    End Function



    <WebMethod()> _
    Public Function ZamowienieDaneDodatkowePobierz(ByVal sesja As Byte(), ByVal zamowienieId As Integer) As ZamowienieDaneDodatkoweWynik

        Dim wynik As New ZamowienieDaneDodatkoweWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_DANE_DODATKOWE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieTypListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_TYP_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieListPrzewozowyZmienZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal numerListu As String) As ZamowienieListPrzewozowyZmienZapiszWynik

        Dim wynik As New ZamowienieListPrzewozowyZmienZapiszWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", NZ(projektId, DBNull.Value))
            htParams.Add("@ZAMOWIENIE_ID", NZ(zamowienieId, DBNull.Value))
            htParams.Add("@LIST_NR", NZ(numerListu, DBNull.Value))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_LIST_PRZEWOZOWY_ZMIEN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function


    <WebMethod()> _
    Public Function ZamowienieParagonBlednyDanePobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer) As ZamowienieParagonBlednyDanePobierzWynik

        Dim wynik As New ZamowienieParagonBlednyDanePobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_PARAGON_BLEDNE_DANE", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ZamowienieParagonBlednyDanePoprawZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal paragonId As Integer, ByVal drukarkaId As Integer, _
                                                            ByVal paragonHeader As String, ByVal paragonNumer As String, ByVal kwotaVatA As String, ByVal kwotaVatB As String, _
                                                            ByVal kwotaVatC As String, ByVal kwotaVatD As String, ByVal kwotaSumaBrutto As String, _
                                                            ByVal skanParagonu As Byte(), ByVal paragonPoprawny As Boolean) As ZamowienieParagonBlednyDanePoprawZapiszWynik

        Dim wynik As New ZamowienieParagonBlednyDanePoprawZapiszWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", NZ(projektId, DBNull.Value))
            htParams.Add("@ZAMOWIENIE_ID_IN", NZ(zamowienieId, DBNull.Value))
            htParams.Add("@PARAGON_ID_IN", NZ(paragonId, DBNull.Value))
            htParams.Add("@DRUKARKA_ID_IN", NZ(drukarkaId, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_HEADER_IN", NZ(paragonHeader, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_NR_IN", NZ(paragonNumer, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_VAT_A_IN", NZ(kwotaVatA, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_VAT_B_IN", NZ(kwotaVatB, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_VAT_C_IN", NZ(kwotaVatC, DBNull.Value))
            htParams.Add("@DRUKARKA_PARAGON_VAT_D_IN", NZ(kwotaVatD, DBNull.Value))
            htParams.Add("@WARTOSC_PARAGONU_BRUTTO", NZ(kwotaSumaBrutto, DBNull.Value))
            htParams.Add("@PLIK_SKAN", NZ(skanParagonu, DBNull.Value))
            htParams.Add("@POPRAWNY_PARAGON", NZ(paragonPoprawny, DBNull.Value))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_PARAGON_BLEDNE_DANE_POPRAW", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowinenieAtrybutyEdycjaZapisz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zamowienieId As Integer, ByVal atrybutyDS As DataSet) As ZamowinenieAtrybutyEdycjaZapiszWynik

        Dim wynik As New ZamowinenieAtrybutyEdycjaZapiszWynik

        If Not atrybutyDS Is Nothing AndAlso atrybutyDS.Tables.Count > 0 Then

            Dim atrybutyDt As DataTable = atrybutyDS.Tables(0)
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projekt_id)
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@TABLEATRYBUTYZAMOWIENIA_IN", atrybutyDt)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_ATRYBUT_EDYTUJ_ZAPISZ", htParams)
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane listy zamówień do zapisania"
        End If


        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowinenieFvDoParagonuGeneruj(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal dataWystawienia As DateTime) As ZamowinenieFvDoParagonuGenerujWynik

        Dim wynik As New ZamowinenieFvDoParagonuGenerujWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@DATA_WYSTAWIENIA", IIf(dataWystawienia = DateTime.MinValue, DBNull.Value, dataWystawienia))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_WYGENERUJ_FV_DO_PARAGONU", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


    <WebMethod()> _
    Public Function ZamowinenieUwagiEdycjaZapisz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zamowienieId As Integer, ByVal uwagi As String) As ZamowinenieUwagiEdycjaZapiszWynik

        Dim wynik As New ZamowinenieUwagiEdycjaZapiszWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@UWAGA", uwagi)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_UWAGI_EDYTUJ_ZAPISZ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowieniePunktOdbioruZmienZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal pickupPoint As String, ByVal pickupPointAlt As String) As ZamowieniePunktOdbioruZmienZapiszWynik

        Dim wynik As New ZamowieniePunktOdbioruZmienZapiszWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", NZ(projektId, DBNull.Value))
            htParams.Add("@ZAMOWIENIE_ID", NZ(zamowienieId, DBNull.Value))
            htParams.Add("@PUNKT_ODBIORU_NR_IN", NZ(pickupPoint, DBNull.Value))
            htParams.Add("@PUNKT_ODBIORU_ALT_NR_IN", NZ(pickupPointAlt, DBNull.Value))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_PUNKT_ODBIORU_ZMIEN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieOdblokujDoRealizacji(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieNr As String, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZAMOWIENIE_NR_IN", zamowienieNr)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_ODBLOKUJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZamowienieOznaczOczekujeNaStan(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieNr As String, ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZAMOWIENIE_NR_IN", zamowienieNr)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_OCZEKUJE_NA_STAN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


    <WebMethod()> _
    Public Function ZamowienieFakturaEnovaZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal zamowienieNumer As String, ByVal plikDane As Byte(), ByVal numer As String, ByVal data As DateTime, ByVal nazwa As String, ByVal wartosc As Decimal) As ZamowienieFakturaEnovaZapiszWynik

        Dim wynik As New ZamowienieFakturaEnovaZapiszWynik

        Using cnn As New ConnectionSuperPaker

            Dim idOut As Integer

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", NZ(projektId, DBNull.Value))
            htParams.Add("@ZAMOWIENIE_ID_IN", NZ(zamowienieId, DBNull.Value))
            htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", NZ(zamowienieNumer, DBNull.Value))
            htParams.Add("@PLIK_IN", NZ(plikDane, DBNull.Value))
            htParams.Add("@DOK_NR", NZ(numer, DBNull.Value))
            htParams.Add("@DOK_DATA", NZ(data, DBNull.Value))
            htParams.Add("@DOK_NAZWA", NZ(nazwa, DBNull.Value))
            htParams.Add("@DOK_WARTOSC_BRUTTO", NZ(wartosc, DBNull.Value))
            htParams.Add("@OUT_ZEWNETRZNY_ID", NZ(idOut, DBNull.Value))


            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_ENOVA_FV_ZEWNETRZNY_ZAPISZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If procRes.Status = 0 Then
                    wynik.id = NZ(htParams("@OUT_ZEWNETRZNY_ID"), -1)
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()>
    Public Function ZamowienieDokumentZewnetrznyDaneZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal zamowienieNumer As String, ByVal plikDane As Byte(), ByVal numer As String, ByVal data As DateTime, ByVal nazwa As String, ByVal typDokumentuId As Integer) As ZamowienieDokumentZewnetrznyDaneZapiszWynik

        Dim wynik As New ZamowienieDokumentZewnetrznyDaneZapiszWynik

        Using cnn As New ConnectionSuperPaker

            Dim idOut As Integer

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", NZ(projektId, DBNull.Value))
            htParams.Add("@ZAMOWIENIE_ID_IN", NZ(zamowienieId, DBNull.Value))
            htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", NZ(zamowienieNumer, DBNull.Value))
            If IsNothing(plikDane) = False Then
                htParams.Add("@PLIK_IN", plikDane)
            End If

            htParams.Add("@DOK_NR", NZ(numer, DBNull.Value))
            htParams.Add("@DOK_DATA", NZ(data, DBNull.Value))
            htParams.Add("@DOK_NAZWA", NZ(nazwa, DBNull.Value))
            htParams.Add("@TYP_DOKUMENTU", NZ(typDokumentuId, DBNull.Value))
            htParams.Add("@OUT_ZEWNETRZNY_ID", NZ(idOut, DBNull.Value))


            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_ZEWNETRZNY_DANE_ZAPISZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If procRes.Status = 0 Then
                    wynik.id = NZ(htParams("@OUT_ZEWNETRZNY_ID"), -1)
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

#End Region

#Region "Zamówienia import ręczny"

    <WebMethod()> _
    Public Function ImpDostawaMetodaListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_DOSTAWA_METODA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamTypListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ZAMOWIENIE_TYP_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamStatusImportuListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ZAMOWIENIE_STATUS_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamJednostkaMiaryListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_JEDNOSTKA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpKontrahentListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                           ByVal dane As String) As ImpKontrahentListaWynik

        Dim wynik As New ImpKontrahentListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@WYSZUKAJ_DANE_IN", dane)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_KONTRAHENT_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpIdKlientaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As String) As ImpKontrahentIdKlientaPobierzWynik
        Dim wynik As New ImpKontrahentIdKlientaPobierzWynik
        Dim klientIDOUT As String = ""

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@WYSZUKAJ_DANE_IN", dane)
            htParams.Add("@OUT_KLIENT_ID:25", klientIDOUT)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_IMP_KONTRAHENT_KLIENT_ID_POBIERZ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status > -1 Then
                    wynik.klientID = NZ(htParams("@OUT_KLIENT_ID:25"), "")
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpAdresListaPobierz(ByVal sesja As Byte(), projektId As Integer,
                                           ByVal dane As String) As AdresDlaKoduListaWynik

        Dim wynik As New AdresDlaKoduListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@WYSZUKAJ_DANE_IN", dane)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ADRES_DLA_KODU_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpPakietEdytujZapisz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal impPakietID As Integer,
                                          ByVal impPakietTypID As Integer, ByVal pakietNazwa As String) As StatusWynik
        Dim wynik As New StatusWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            htParams.Add("@IMP_PAKIET_ID_IN", impPakietID)
            htParams.Add("@IMP_TYP_ID_IN", impPakietTypID)
            htParams.Add("@IMP_NAZWA_IN", pakietNazwa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_IMP_PAKIET_DODAJ_ZMIEN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpPakietListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime, ByVal wyszukajTekst As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajTekst)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_IMP_PAKIET_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ImpPlatnoscMetodaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_PLATNOSC_METODA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowienieDodajEdytuj(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal impZamowienieID As Integer,
                                          ByVal zamowienieDS As DataSet, ByVal zamowieniePozycjeDS As DataSet, ByVal zamowienieAtrybutyDS As DataSet) As ImpZamowienieZapiszWynik
        Dim wynik As New ImpZamowienieZapiszWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim zamowienieDT As DataTable
        Dim zamowieniePozycjeDT As DataTable
        Dim zamowienieAtrybutyDT As DataTable
        If zamowienieDS.Tables.Count > 0 And zamowieniePozycjeDS.Tables.Count > 0 And zamowienieAtrybutyDS.Tables.Count > 0 Then
            zamowienieDT = zamowienieDS.Tables(0).Copy
            zamowieniePozycjeDT = zamowieniePozycjeDS.Tables(0).Copy
            zamowienieAtrybutyDT = zamowienieAtrybutyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektID)
                htParams.Add("@ZAMOWIENIE_ID_IN", impZamowienieID)

                htParams.Add("@ZAMOWIENIE_POZYCJE_LISTA_IN", zamowieniePozycjeDT)

                htParams.Add("@ZAMOWIENIE_DANE_LISTA_IN", zamowienieDT)
                htParams.Add("@ZAMOWIENIE_ATRYBUT_LISTA_IN", zamowienieAtrybutyDT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ZAMOWIENIE_DODAJ_ZMIEN", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.zamowienie = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status_opis = "Nieprawidłowe dane listy zamówień do zapisania"

            If zamowienieDS.Tables.Count = 0 Then wynik.status_opis = "Nie przesłano danych zamówienia"
            If zamowieniePozycjeDS.Tables.Count = 0 Then wynik.status_opis = "Nie przesłano pozycji zamówienia"
            If zamowienieAtrybutyDS.Tables.Count = 0 Then wynik.status_opis = "Nie przesłano atrybutów zamówienia"

            wynik.zamowienie = Nothing
            wynik.status = -1

        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowienieListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal dane As String) As ImpZamowienieListaWynik
        Dim wynik As New ImpZamowienieListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@WYSZUKAJ_DANE_IN", dane)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_IMP_ZAMOWIENIE_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ImpZamowieniePozycjaListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal zamowienieId As Integer) As ImpZamowieniePozycjaListaWynik
        Dim wynik As New ImpZamowieniePozycjaListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ZAMOWIENIE_POZYCJA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ImpKontrahentDodajEdytuj(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal impkontrahentTypID As Integer, ByVal impKontrahentID As Integer, ByVal impKontrahentDepartamentID As Integer,
                                          ByVal kontrahentDS As DataSet) As ImpKontrahentZapiszWynik
        Dim wynik As New ImpKontrahentZapiszWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim daneZamowienia As DataTable

        If kontrahentDS.Tables.Count > 0 Then
            daneZamowienia = kontrahentDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektID)
                htParams.Add("@KONTRAHENT_ID_IN", impKontrahentID)
                htParams.Add("@KONTRAHENT_TYP_ID_IN", impkontrahentTypID)
                htParams.Add("@KONTRAHENT_DEPARTMENT_ID_IN", impKontrahentDepartamentID)
                htParams.Add("@ZAMOWIENIE_DANE_LISTA_IN", daneZamowienia)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_KONTRAHENT_DODAJ_ZMIEN", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane zamówienia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowienieUsun(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal zamowienieID As Integer) As ImpZamowienieUsunWynik
        Dim wynik As New ImpZamowienieUsunWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieID)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_IMP_ZAMOWIENIE_USUN", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowieniePakietUsun(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal pakietID As Integer) As ImpPakietUsunWynik
        Dim wynik As New ImpPakietUsunWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            htParams.Add("@IMP_PAKIET_ID_IN", pakietID)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_IMP_PAKIET_USUN", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowienieUstawStatus(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal zamowieniaDS As DataSet, ByVal nowyStatusID As Integer) As ImpZamowienieUstawStatusWynik
        Dim wynik As New ImpZamowienieUstawStatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableZamowienia As DataTable = New DataTable

        If zamowieniaDS.Tables.Count > 0 Then
            dTableZamowienia = zamowieniaDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektID)
                htParams.Add("@IMP_ZAMOWIENIE_LISTA_IN", dTableZamowienia)
                htParams.Add("@STATUS_ID_IN", nowyStatusID)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_IMP_ZAMOWIENIE_LISTA_STATUS_USTAW", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane listy zamówień do zmiany statusu"
        End If

        Return wynik
    End Function



    <WebMethod()> _
    Public Function ImpZamowieniePrzeniesDoSTG(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal zamowieniaDS As DataSet) As ImpZamowieniePrzeniesDoSTGWynik
        Dim wynik As New ImpZamowieniePrzeniesDoSTGWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableZamowienia As DataTable = New DataTable

        If zamowieniaDS.Tables.Count > 0 Then
            dTableZamowienia = zamowieniaDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektID)
                htParams.Add("@IMP_ZAMOWIENIE_LISTA_IN", dTableZamowienia)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "STG_IMP_ZAMOWIENIE_STG_ZAPISZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane listy zamówień do zatwierdzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ImpZamowienieDzialListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer
                                                             ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_IMP_ZAMOWIENIE_DZIAL_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


    <WebMethod()> _
    Public Function ImpDzialEdytujZapisz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal dzialID As Integer,
                                          ByVal dzialTypID As Integer, ByVal dzialKlientId As String, ByVal dzialNazwa As String) As StatusWynik
        Dim wynik As New StatusWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            htParams.Add("@IMP_DZIAL_ID_IN", dzialID)
            htParams.Add("@IMP_DZIAL_TYP_ID_IN", dzialTypID)
            htParams.Add("@IMP_DZIAL_KLIENT_ID_IN", dzialKlientId)
            htParams.Add("@IMP_DZIAL_NAZWA_IN", dzialNazwa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_IMP_DZIAL_DODAJ_ZMIEN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DostawaGabarytyListaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal kurierId As Integer?, ByVal gabarytTypId As Integer?) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            If kurierId <= 0 Then kurierId = Nothing
            If gabarytTypId <= 0 Then gabarytTypId = Nothing
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@KURIER_ID", NZ(kurierId, DBNull.Value))
            htParams.Add("@PACZKA_GABARYT_TYP_ID", NZ(gabarytTypId, DBNull.Value))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_KURIER_PRZESYLKA_GABARYT_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


#End Region

#Region "Funkcje"

    <WebMethod()> _
    Public Function FunkcjeListaPobierz(ByVal uzytkownik_Id As Integer, ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownik_Id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_FUNKCJA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KontrolkiUprawnieniaPobierz(ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_UZYTKOWNIK_KONTR_UPRAW_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Blokada"

    <WebMethod()> _
    Public Function BlokadaUsun(ByVal sesja As Byte(), ByVal zablokowaneID As Integer, ByVal Tabela As String) As BlokadaUsunWynik

        Dim wynik As New BlokadaUsunWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ID_BLOKOWANE", zablokowaneID)
            htParams.Add("@TABELA", Tabela)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_TABELA_ODBLOKUJ", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function BlokadyUzytkownikaWszystkieUsun(ByVal sesja As Byte()) As BlokadaUsunWynik

        Dim wynik As New BlokadaUsunWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_TABELA_ODBLOKUJ_LOGOUT", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using
        Return wynik
    End Function

    <WebMethod()> _
    Public Function BlokadaInfoPobierz(ByVal sesja As Byte(), ByVal zablokowaneID As Integer, ByVal Tabela As String) As BlokadaWynik

        Dim wynik As New BlokadaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ID_BLOKOWANE_IN", zablokowaneID)
            htParams.Add("@TABELA_IN", Tabela)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_TABELA_BLOKADA_INFO_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function BlokadaUsunAwaryjnie(ByVal sesja As Byte(), ByVal zablokowaneID As Integer, ByVal uzytkownikID As Integer, ByVal Tabela As String) As BlokadaUsunWynik
        Dim cnn As SqlConnection
        Dim wynik As New BlokadaUsunWynik

        'łączymy do bazyO
        Try
            cnn = New SqlConnection()
            cnn.ConnectionString = connectionString
            cnn.Open()
        Catch ex As Exception
            ' logger.Error("UserEdytujAnuluj:Błąd komunikacji z bazą: ", ex)
            Return wynik
        End Try

        Dim cmd As New SqlClient.SqlCommand("UP_TABELA_ODBLOKUJ_AWARYJNIE", cnn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 1
        cmd.Parameters.AddWithValue("@sesja", sesja)
        cmd.Parameters.AddWithValue("@ID_BLOKOWANE", zablokowaneID)
        cmd.Parameters.AddWithValue("@ID_UZYTKOWNIKA", uzytkownikID)
        cmd.Parameters.AddWithValue("@TABELA", Tabela)
        cmd.Parameters.Add("@status", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@status_opis", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd komunikacji z bazą: " & ex.Message & kontaktIt
            cnn.Close()
            Return wynik
        End Try

        wynik.status = cmd.Parameters("@status").Value
        wynik.status_opis = cmd.Parameters("@status_opis").Value
        cnn.Close()
        Return wynik

    End Function

#End Region

#Region "Funkcje Pomocnicze"

    Public Shared Function NZ(ByVal S As Object, ByVal Def As Object) As Object
        If IsDBNull(S) Then
            Return Def
        Else
            If Not (S Is Nothing) Then
                Return (S)
            Else
                Return Def
            End If
        End If
    End Function

#End Region

#Region "Awiza"

    <WebMethod()> _
    Public Function AwizaListaPobierz(ByVal sesja As Byte(), ByVal uzytkownik_id As Integer, ByVal projekt_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownik_id)
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_AWIZO_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function OsobyKontaktoweListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_OSOBY_KONTAKTOWE_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function OsobyKontaktoweDodaj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal Osoba As String, ByVal telefon As String) As OsobyKontaktoweDodajWynik

        Dim wynik As New OsobyKontaktoweDodajWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            Dim osobaIdOUT As Integer = 0
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@OSOBA_NAZWA_IN", Osoba)
            htParams.Add("@TELEFON_IN", telefon)
            htParams.Add("@OUT_OSOBA_ID", osobaIdOUT)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_AWIZO_OSOBY_KONTAKTOWE_DODAJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If procRes.Status = 0 Then
                    wynik.osobaId = NZ(htParams("@OUT_OSOBA_ID"), "")
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function AwizoEdytuj(ByVal sesja As Byte(), ByVal AwizoEdytowaneId As Integer, projekt_id As Integer) As AwizoEdytujWynik

        Dim wynik As New AwizoEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim projektPrefixOUT As String = ""
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@AWIZO_ID_EDYTOWANE_IN", AwizoEdytowaneId)
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@OUT_PROJEKT_PREFIX:25", projektPrefixOUT)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_AWIZO_EDYCJA_ROZPOCZNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.project_prefix = NZ(htParams("@OUT_PROJEKT_PREFIX:25"), "")
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If

            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function AwizoZatwierdzIska(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet) As AwizoZapiszWynik
        Dim wynik As New AwizoZapiszWynik
        wynik.status = -1
        wynik.status_opis = "Metoda: AwizoZatwierdzIska - błąd krytyczny. Skontaktuj się z IT"
        'Zawsze zapisuje awizo aby znać jego ID
        'Pytam baze o dane produktów, login, hasło i URL
        'Otwieram order w iska i dodaje produktu
        'Jeśli produkty dodały się to zatwierdzam u nas zamówienie
        'Jeśli zamówienie zatwierdziło się to zatwierdzam zamówienie w ISKA
        'Notuję wynik akceptacji ISKA i niezależnie od efektu zapisuję ten wynik w bazie SP

        Dim wynikAwizoZapisz As New AwizoZapiszWynik
        Try
            wynikAwizoZapisz = AwizoEdytujZapisz(sesja, dane, 0)
            wynik.status = wynikAwizoZapisz.status
            wynik.status_opis = wynikAwizoZapisz.status_opis
            wynik.awizoId = wynikAwizoZapisz.awizoId
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception zapis awiza: " + ex.Message
            Return wynikAwizoZapisz
        End Try

        'Zwracam wynik gdy się nie zapisało
        If Not wynik.status = 0 Then Return wynik

        Dim wynikIskaPozycje As AwizoIskaPozycjeWyslijWynik = AwizoIskaPozycjeWyslij(sesja, projektId, wynikAwizoZapisz.awizoId)

        wynik.status = wynikIskaPozycje.status
        wynik.status_opis = wynikIskaPozycje.status_opis
        If Not wynik.status = 0 Then Return wynik

        Try
            dane.Tables(0).Rows(0)("AWIZO_ID") = wynik.awizoId 'bez tego tworzyło się drugie awizo, zamiast zatwierdzic poprzednie
            wynik = AwizoEdytujZapisz(sesja, dane, 1)
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception zatwierdzanie awiza: " + ex.Message
            Return wynik
        End Try
        If Not wynik.status = 0 Then Return wynik

        Dim wynikIskaPotwierdzenie As AwizoIskaPotwierdzenieZamknijIZapiszWynik
        Try
            wynikIskaPotwierdzenie = AwizoIskaPotwierdzenieZamknijIZapisz(sesja, projektId, wynikAwizoZapisz.awizoId, wynikIskaPozycje.awizoId, wynikIskaPozycje.user, wynikIskaPozycje.password, wynikIskaPozycje.url)
            wynik.status = wynikIskaPotwierdzenie.status
            wynik.status_opis = wynikIskaPotwierdzenie.status_opis
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception zatwierdzanie zapis statusu zatwierdzania ISKA: " + ex.Message
        End Try


        Return wynik
    End Function
    <WebMethod()> _
    Public Function AwizoIskaWyslijPonownie(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal awizoId As Integer) As AwizoIskaWyslijPonownieWynik
        Dim wynik As New AwizoIskaWyslijPonownieWynik
        wynik.status = -1
        wynik.status_opis = "Metoda: AwizoIskaWyslijPonownie - błąd krytyczny. Skontaktuj się z IT"
        'Pytam baze o dane produktów, login, hasło i URL
        'Otwieram order w iska i dodaje produktu
        'Jeśli produkty dodały się to zatwierdzam ISKA
        'Notuję wynik akceptacji ISKA i niezależnie od efektu zapisuję ten wynik w bazie SP

        Dim wynikIskaPozycje As AwizoIskaPozycjeWyslijWynik = AwizoIskaPozycjeWyslij(sesja, projektId, awizoId)

        wynik.status = wynikIskaPozycje.status
        wynik.status_opis = wynikIskaPozycje.status_opis
        If Not wynik.status = 0 Then Return wynik

        Dim wynikIskaPotwierdzenie As AwizoIskaPotwierdzenieZamknijIZapiszWynik
        Try
            wynikIskaPotwierdzenie = AwizoIskaPotwierdzenieZamknijIZapisz(sesja, projektId, awizoId, wynikIskaPozycje.awizoId, wynikIskaPozycje.user, wynikIskaPozycje.password, wynikIskaPozycje.url)
            wynik.status = wynikIskaPotwierdzenie.status
            wynik.status_opis = wynikIskaPotwierdzenie.status_opis
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception zatwierdzanie zapis statusu zatwierdzania ISKA: " + ex.Message
        End Try

        Return wynik
    End Function
#Region "ISKA funkcje pomocnicze"
    'nie webmethod - funkcja pomocnicza dla ISKA
    Private Function AwizoIskaPozycjeWyslij(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal awizoId As Integer) As AwizoIskaPozycjeWyslijWynik

        Dim wynik As New AwizoIskaPozycjeWyslijWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@AWIZO_ID", awizoId)
            htParams.Add("@PROJEKT_ID", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ISKA_AWIZO_POZYCJE", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        'Wysyłanie danych do ISKA
        Dim produktyIska As New DataTable
        Dim ustawieniaIska As New DataTable


        Try
            If wynik.status = 0 Then
                If wynik.dane.Tables.Count > 1 Then
                    produktyIska = wynik.dane.Tables(0)
                    ustawieniaIska = wynik.dane.Tables(1)
                Else
                    wynik.status = -1
                    wynik.status_opis = "AwizoIskaPozycjePobierz - zwrócono za małą liczbę tabel"
                End If
            End If

            If produktyIska.Rows.Count = 0 Then
                wynik.status = -1
                wynik.status_opis = "Brak produktów do wysłania do ISKA"
            End If

            If ustawieniaIska.Rows.Count = 0 Then
                wynik.status = -1
                wynik.status_opis = "Brak ustawień ISKA"
            End If
            'Tabela 0
            'awizo_id
            'awizo_pozycja_id
            'ILOSC_AWIZOWANA
            'EAN
            'KLIENT_PRODUKT_NR
            'bloz
            'awizo_iska_status_nazwa
            'AWIZO_ISKA_STATUS_ID

            'Tabela 1
            'iska_login
            'iska_haslo
            'iska_url
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception pobieranie ustawień ISKA z SP: " + ex.Message
            Return wynik
        End Try

        If Not wynik.status = 0 Then Return wynik

        wynik.user = NZ(ustawieniaIska.Rows(0)("iska_login"), "")
        wynik.password = NZ(ustawieniaIska.Rows(0)("iska_haslo"), "")
        wynik.url = NZ(ustawieniaIska.Rows(0)("iska_url"), "")

        If wynik.user = "" OrElse wynik.password = "" OrElse wynik.url = "" Then
            wynik.status = -1
            wynik.status_opis = "Dane logowania do ISKA są puste"
            Return wynik
        End If

        Dim wsIska As New pl.net.iska.ecommerce_Model_ApiService
        wsIska.Url = wynik.url.Replace("?wsdl", "")
        Dim cred = New System.Net.NetworkCredential(wynik.user, wynik.password)
        Dim credentials = cred.GetCredential(New Uri(wsIska.Url), "Basic")
        wsIska.Credentials = credentials
        wsIska.PreAuthenticate = True

        Dim supplyOppened As Boolean
        Try
            supplyOppened = wsIska.supplyOpen(awizoId)
            wynik.awizoId = awizoId
            If supplyOppened = False Then
                wynik.status = -1
                wynik.status_opis = "ISKA - nie udało się otworzyć zamówienia"
                Return wynik
            End If

        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception ISKA supplyOpen: " + ex.Message
            Return wynik
        End Try

        Try
            For Each produktRow As DataRow In produktyIska.Rows
                'If wsIska.orderSetProduct(produktRow("KLIENT_PRODUKT_NR"), orderId, produktRow("bloz"), produktRow("EAN"), produktRow("ILOSC_AWIZOWANA")) = False Then
                If wsIska.supplySetProduct(produktRow("awizo_pozycja_id"), awizoId, produktRow("KLIENT_PRODUKT_NR"), produktRow("bloz"), produktRow("EAN"), produktRow("ILOSC_AWIZOWANA")) = False Then
                    wynik.status = -1
                    wynik.status_opis = "Nie udało się dodać produktu do zamówienia ISKA"
                    Return wynik
                End If
            Next
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception ISKA supplySetProduct: " + ex.Message
            Return wynik
        End Try

        Return wynik
    End Function
    'nie webmethod - funkcja pomocnicza dla ISKA
    Private Function AwizoIskaPotwierdzenieZamknijIZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal awizoId As Integer, ByVal orderId As Integer, ByVal user As String, ByVal password As String, ByVal urlIska As String) As AwizoIskaPotwierdzenieZamknijIZapiszWynik
        Dim wynik As New AwizoIskaPotwierdzenieZamknijIZapiszWynik

        Dim wsIska As New pl.net.iska.ecommerce_Model_ApiService
        wsIska.Url = urlIska.Replace("?wsdl", "")
        Dim cred = New System.Net.NetworkCredential(user, password)
        Dim credentials = cred.GetCredential(New Uri(wsIska.Url), "Basic")
        wsIska.Credentials = credentials
        wsIska.PreAuthenticate = True

        Dim czyZatwierdzone As Boolean = False
        Try
            czyZatwierdzone = wsIska.supplyClose(awizoId)
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Exception zatwierdzanie ISKA orderClose: " + ex.Message
            'nie zwracam bo już zatwierdziłem u nas
        End Try

        Dim dSet As DataSet = New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@AWIZO_ID", awizoId)
            htParams.Add("@ORDERID", orderId)
            htParams.Add("@WYNIK", czyZatwierdzone)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ISKA_POTWIERDZENIA", htParams, dSet)
                'wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
#End Region

    <WebMethod()> _
    Public Function AwizoEdytujZapisz(ByVal sesja As Byte(), ByVal dane As DataSet, zatwierdz As Integer) As AwizoZapiszWynik

        Dim wynik As New AwizoZapiszWynik

        If dane.Tables.Count > 1 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim AwizoTable As New DataTable
            Dim AwizoPozycjeTable As New DataTable
            Dim awizoIdEtytowaneOUT As Integer = 0

            If dane.Tables.Count > 1 Then
                AwizoTable = dane.Tables(0).Copy
                AwizoPozycjeTable = dane.Tables(1).Copy

                'TO TYP TABLICOWY. OSTATNIO DODANA BYŁA KOLUMNA MAGAZYN_WIRTUALNY_ID - STĄD TEN DODATEK
                If AwizoTable.Columns.Contains("MAGAZYN_WIRTUALNY_ID") = False Then
                    AwizoTable.Columns.Add("MAGAZYN_WIRTUALNY_ID", GetType(Integer))
                    AwizoTable.AcceptChanges()
                End If

                 If AwizoTable.Columns.Contains("PACZKI") = False Then
                    AwizoTable.Columns.Add("PACZKI", GetType(Integer))
                    AwizoTable.AcceptChanges()
                End If

				If AwizoTable.Columns.Contains("PALETY") = False Then
                    AwizoTable.Columns.Add("PALETY", GetType(Integer))
                    AwizoTable.AcceptChanges()
                End If

                'Nowe kolumny w typie tablicowym.
                If AwizoPozycjeTable.Columns.Contains("PARTIA") = False Then
                    AwizoPozycjeTable.Columns.Add("PARTIA", GetType(String))
                    AwizoPozycjeTable.AcceptChanges()
                End If
                If AwizoPozycjeTable.Columns.Contains("DATA_WAZNOSCI") = False Then
                    AwizoPozycjeTable.Columns.Add("DATA_WAZNOSCI", GetType(DateTime))
                    AwizoPozycjeTable.AcceptChanges()
                End If

            End If

            'Poprawka dat
            For Each dRow In AwizoPozycjeTable.Rows
                If CDate(NZ(dRow("DATA_WAZNOSCI"), Date.MinValue)) = Date.MinValue Then
                    dRow("DATA_WAZNOSCI") = DBNull.Value
                End If
            Next


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@TableAwizo_in", AwizoTable)
                htParams.Add("@TableAwizoPozycje_in", AwizoPozycjeTable)
                htParams.Add("@AWIZO_ZATWIERDZ_IN", zatwierdz)
                htParams.Add("@OUT_AWIZO_ID_EDYTOWANE", awizoIdEtytowaneOUT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_AWIZO_EDYCJA_ZAPISZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.awizoId = NZ(htParams("@OUT_AWIZO_ID_EDYTOWANE"), 0)
                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function AwizoAnuluj(ByVal sesja As Byte(), ByVal awizoId As Integer, ByVal projektId As Integer) As StatusWynik

        Dim wynik As New StatusWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@AWIZO_ID_ANULOWANE_IN", awizoId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_AWIZO_USUN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()> _
    Public Function AwizoDostawaZrodloListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As AwizoDostawaZrodloListaWynik
        Dim wynik As New AwizoDostawaZrodloListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim mCount As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_AWIZO_DOSTAWA_ZRODLO_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function MagazynyWirtualneListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal tylkoAktywne As Nullable(Of Boolean)) As MagazynyWirtualneListaPobierzWynik
        Dim wynik As New MagazynyWirtualneListaPobierzWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@AKTYWNY", IIf(tylkoAktywne Is Nothing, DBNull.Value, tylkoAktywne))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_MAGAZYN_WIRTUALNY_LISTA", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function AwizoStatusListaPobierz(ByVal sesja As Byte(), projektId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_AWIZO_STATUSY_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function AwizoStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer, ByVal uzytkownikId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String, ByVal dataOd As DateTime, ByVal dataDo As DateTime,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer,
                                      ByVal sortPo As String,
                                      ByVal sortAsc As Boolean,
                                      ByVal statusyDS As DataSet) As AwizoStronaWynik

        Dim wynik As New AwizoStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0
        Dim dTableStatusy As DataTable = New DataTable

        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    htParams.Add("@UZYTKOWNIK_ID_IN", uzytkownikId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STRONA_NR_IN", stronaNumer)
                    htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                    htParams.Add("@DATA_OD", IIf(dataOd = DateTime.MinValue, DBNull.Value, dataOd))
                    htParams.Add("@DATA_DO", IIf(dataDo = DateTime.MinValue, DBNull.Value, dataDo))
                    htParams.Add("@SORT_PO_IN", sortPo)
                    htParams.Add("@SORT_KIERUNEK_IN", sortAsc)
                    htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_AWIZO_LISTA_POBIERZ_STRONA", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    If procRes.Status = 0 Then
                        wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                    End If
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If
        Return wynik

    End Function

#End Region

#Region "Dostawcy"
    <WebMethod()> _
    Public Function DostawcyListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_DOSTAWCY_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DostawcaEdytuj(ByVal sesja As Byte(), ByVal DostawcaEdytowanyId As Integer, projekt_id As Integer) As DostawcaEdytujWynik

        Dim wynik As New DostawcaEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim KlientIdOUT As Integer = -1

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@DOSTAWCA_ID_EDYTOWANE_IN", DostawcaEdytowanyId)
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@OUT_KLIENT_ID", KlientIdOUT)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_DOSTAWCA_EDYCJA_ROZPOCZNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then

                    wynik.klient_id = NZ(htParams("@OUT_KLIENT_ID"), 0)
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If

            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DostawcyEdytujZapisz(ByVal sesja As Byte(), ByVal dtDostawcy As DataTable) As DostawcyZapiszWynik

        Dim wynik As New DostawcyZapiszWynik

        If dtDostawcy.Rows.Count > 0 Then


            Dim awizoIdEtytowaneOUT As Integer = 0



            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@TableDostawcy_in", dtDostawcy)

                ' htParams.Add("@AWIZO_ZATWIERDZ_IN", zatwierdz)
                htParams.Add("@OUT_DOSTAWCA_ID_EDYTOWANE", awizoIdEtytowaneOUT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOSTAWCY_EDYCJA_ZAPISZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.dostawcaId = NZ(htParams("@OUT_DOSTAWCA_ID_EDYTOWANE"), 0)

                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

#End Region

#Region "Kompletacja"
    <WebMethod()> _
    Public Function ZamowienieKompletacjaListaFiltrListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_FILTR_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieKompletacjaMasowaListaPobierz(ByVal sesja As Byte(), projektId As Integer, wartosc As String, typDokumentu As Integer, ByVal dane As DataSet
                                           ) As ZamowienieKompletacjaMasowaListaWynik

        Dim wynik As New ZamowienieKompletacjaMasowaListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim KompletacjaElementy As New DataTable

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            If dane.Tables.Count > 0 Then
                KompletacjaElementy = dane.Tables(0).Copy
            End If

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@WARTOSC_IN", wartosc)
            htParams.Add("@TYP_DOKUMENTU_ID_IN", typDokumentu)
            htParams.Add("@TABLE_ELEMENTY_LISTA", KompletacjaElementy)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_MASOWA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieKompletacjaListaPobierz(ByVal sesja As Byte(),
                                         ByVal projektNazwa As String,
                                      ByVal schematNazwa As String,
                                      ByVal zamowienieTypId As Integer,
                                      ByVal aktualnyEtapId As Integer,
                                      ByVal qguarZL As String,
                                      ByVal klientNumerZamowienia As String,
                                      ByVal dataPobraniaOdKlienta As DateTime?,
                                      ByVal dataPobraniaDoPakowania As DateTime?,
                                      ByVal dataSpakowania As DateTime?,
                                      ByVal pakujacyId As Integer,
                                      ByVal przydzielajacyId As Integer,
                                      ByVal zamStatusy As DataSet
                                           ) As ZamowienieListaWynik

        Dim wynik As New ZamowienieListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable

        If zamStatusy.Tables.Count > 0 Then
            dTableStatusy = zamStatusy.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_NAZWA_IN", projektNazwa)
                htParams.Add("@SCHEMAT_NAZWA_IN", schematNazwa)
                htParams.Add("@ZAMOWIENIE_TYP_ID_IN", zamowienieTypId)
                htParams.Add("@AKTUALNY_ETAP_ID_IN", aktualnyEtapId)
                htParams.Add("@QGUAR_ZL_IN", qguarZL)
                htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", klientNumerZamowienia)
                htParams.Add("@DATA_POBRANIA_OD_KLIENTA_IN", dataPobraniaOdKlienta)
                htParams.Add("@DATA_POBRANIA_DO_PAKOWANIA_IN", dataPobraniaDoPakowania)
                htParams.Add("@DATA_SPAKOWANIA_IN", dataSpakowania)
                htParams.Add("@PAKUJACY_ID_IN", pakujacyId)
                htParams.Add("@PRZYDZIELAJACY_ID_IN", przydzielajacyId)
                htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
<WebMethod()> _
    Public Function ZamowienieKompletacjaListaPobierzV2(ByVal sesja As Byte(),
                                         ByVal projektNazwa As String,
                                      ByVal schematNazwa As String,
                                      ByVal zamowienieTypId As Integer,
                                      ByVal aktualnyEtapId As Integer,
                                      ByVal qguarZL As String,
                                      ByVal klientNumerZamowienia As String,
                                      ByVal dataPobraniaOdKlienta As DateTime?,
                                      ByVal dataPobraniaDoPakowania As DateTime?,
                                      ByVal dataSpakowania As DateTime?,
                                      ByVal pakujacyId As Integer,
                                      ByVal przydzielajacyId As Integer,
                                      ByVal zamStatusy As DataSet,
                                       ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String
                                           ) As ZamowienieListaWynik

        Dim wynik As New ZamowienieListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable

        If zamStatusy.Tables.Count > 0 Then
            dTableStatusy = zamStatusy.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_NAZWA_IN", projektNazwa)
                htParams.Add("@SCHEMAT_NAZWA_IN", schematNazwa)
                htParams.Add("@ZAMOWIENIE_TYP_ID_IN", zamowienieTypId)
                htParams.Add("@AKTUALNY_ETAP_ID_IN", aktualnyEtapId)
                htParams.Add("@QGUAR_ZL_IN", qguarZL)
                htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", klientNumerZamowienia)
                htParams.Add("@DATA_POBRANIA_OD_KLIENTA_IN", dataPobraniaOdKlienta)
                htParams.Add("@DATA_POBRANIA_DO_PAKOWANIA_IN", dataPobraniaDoPakowania)
                htParams.Add("@DATA_SPAKOWANIA_IN", dataSpakowania)
                htParams.Add("@PAKUJACY_ID_IN", pakujacyId)
                htParams.Add("@PRZYDZIELAJACY_ID_IN", przydzielajacyId)
                htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)

                htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)


                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ZamowienieKompletacjaStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektNazwa As String,
                                      ByVal schematNazwa As String,
                                      ByVal zamowienieTypId As Integer,
                                      ByVal aktualnyEtapId As Integer,
                                      ByVal qguarZL As String,
                                      ByVal klientNumerZamowienia As String,
                                      ByVal dataPobraniaOdKlienta As DateTime?,
                                      ByVal dataPobraniaDoPakowania As DateTime?,
                                      ByVal dataSpakowania As DateTime?,
                                      ByVal pakujacyId As Integer,
                                      ByVal przydzielajacyId As Integer,
                                      ByVal zamStatusy As DataSet,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer,
                                      ByVal sortPo As String,
                                      ByVal sortAsc As Boolean
                            ) As ZamowienieStronaWynik

        Dim wynik As New ZamowienieStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0

        Dim dTableStatusy As DataTable = New DataTable

        If zamStatusy.Tables.Count > 0 Then
            dTableStatusy = zamStatusy.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_NAZWA_IN", projektNazwa)
                htParams.Add("@SCHEMAT_NAZWA_IN", schematNazwa)
                htParams.Add("@ZAMOWIENIE_TYP_ID_IN", zamowienieTypId)
                htParams.Add("@AKTUALNY_ETAP_ID_IN", aktualnyEtapId)
                htParams.Add("@QGUAR_ZL_IN", qguarZL)
                htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", klientNumerZamowienia)
                htParams.Add("@DATA_POBRANIA_OD_KLIENTA_IN", dataPobraniaOdKlienta)
                htParams.Add("@DATA_POBRANIA_DO_PAKOWANIA_IN", dataPobraniaDoPakowania)
                htParams.Add("@DATA_SPAKOWANIA_IN", dataSpakowania)
                htParams.Add("@PAKUJACY_ID_IN", pakujacyId)
                htParams.Add("@PRZYDZIELAJACY_ID_IN", przydzielajacyId)
                htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                htParams.Add("@STRONA_NR_IN", stronaNumer)
                htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                htParams.Add("@SORT_PO_IN", sortPo)
                htParams.Add("@SORT_KIERUNEK_IN", sortAsc)
                htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_STRONA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                If procRes.Status = 0 Then
                    wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                End If
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieKompletacjaWydrukiHurtListaPobierz(ByVal sesja As Byte(),
                                         ByVal projektNazwa As String,
                                      ByVal schematNazwa As String,
                                      ByVal zamowienieTypId As Integer,
                                      ByVal aktualnyEtapId As Integer,
                                      ByVal qguarZL As String,
                                      ByVal klientNumerZamowienia As String,
                                      ByVal dataPobraniaOdKlienta As DateTime?,
                                      ByVal dataPobraniaDoPakowania As DateTime?,
                                      ByVal dataSpakowania As DateTime?,
                                      ByVal pakujacyId As Integer,
                                      ByVal przydzielajacyId As Integer,
                                      ByVal zamStatusy As DataSet
                                           ) As ZamowienieListaWynik

        Dim wynik As New ZamowienieListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable

        If zamStatusy.Tables.Count > 0 Then
            dTableStatusy = zamStatusy.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_NAZWA_IN", projektNazwa)
                htParams.Add("@SCHEMAT_NAZWA_IN", schematNazwa)
                htParams.Add("@ZAMOWIENIE_TYP_ID_IN", zamowienieTypId)
                htParams.Add("@AKTUALNY_ETAP_ID_IN", aktualnyEtapId)
                htParams.Add("@QGUAR_ZL_IN", qguarZL)
                htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", klientNumerZamowienia)
                htParams.Add("@DATA_POBRANIA_OD_KLIENTA_IN", dataPobraniaOdKlienta)
                htParams.Add("@DATA_POBRANIA_DO_PAKOWANIA_IN", dataPobraniaDoPakowania)
                htParams.Add("@DATA_SPAKOWANIA_IN", dataSpakowania)
                htParams.Add("@PAKUJACY_ID_IN", pakujacyId)
                htParams.Add("@PRZYDZIELAJACY_ID_IN", przydzielajacyId)
                htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_LISTA_DOKUMENTY_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieKompletacjaWydrukiHurtOznaczWydrukowaneListaZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowieniaDs As DataSet
                                           ) As ZamowienieKompletacjaWydrukiHurtOznaczWydrukowaneListaZapiszWynik

        Dim wynik As New ZamowienieKompletacjaWydrukiHurtOznaczWydrukowaneListaZapiszWynik

        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableZamowienia As DataTable = New DataTable

        If zamowieniaDs.Tables.Count > 0 Then
            dTableZamowienia = zamowieniaDs.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@TableZamowieniaDokumenty_in", dTableZamowienia)
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_KOMPLETACJA_DOKUMENTY_WYDRUKOWANE", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function KompletacjaPaczkiPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As KompletacjaPaczkiWynik

        Dim wynik As New KompletacjaPaczkiWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_KOMPLETACJA_PACZKI_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPrzesylkaPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As KompletacjaPaczkiWynik

        Dim wynik As New KompletacjaPaczkiWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_PRZESYLKA_GABARYT_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPunktOdbioruPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As KompletacjaPunktOdbioruWynik

        Dim wynik As New KompletacjaPunktOdbioruWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_PRZESYLKA_PUNKT_ODBIORU_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaUslugiDodatkowePobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As KompletacjaUslugiDodatkowePobierzWynik

        Dim wynik As New KompletacjaUslugiDodatkowePobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_PRZESYLKA_USLUGI_DODATKOWE_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaCrossSellDaneKurierPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As KompletacjaCrossSellDaneKurierPobierzPobierzWynik

        Dim wynik As New KompletacjaCrossSellDaneKurierPobierzPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_PRZESYLKA_CS_DANE_KURIER_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPozycjaDodajZmien(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zamowieniePozycjaId As Integer,
                                              ByVal paczkaId As Integer, ByVal ilosc As Integer) As KompletacjaPaczkiWynik

        Dim wynik As New KompletacjaPaczkiWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim iloscOut As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@ZAMOWIENIE_POZYCJA_ID_IN", zamowieniePozycjaId)
                htParams.Add("@PACZKA_ID_IN", paczkaId)
                htParams.Add("@ILOSC_IN", ilosc)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_KOMPLETACJA_POZYCJA_DODAJ_ZMIEN", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPaczkaWagaZmien(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal paczkaId As Integer, ByVal paczkaWaga As Decimal) As StatusWynik

        Dim wynik As New StatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim iloscOut As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@PACZKA_ID_IN", paczkaId)
                htParams.Add("@PACZKA_WAGA_IN", paczkaWaga)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_KOMPLETACJA_PACZKA_WAGA_ZMIEN", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPaczkaZamknij(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal paczkaId As Integer) As KompletacjaPaczkiWynik

        Dim wynik As New KompletacjaPaczkiWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim iloscOut As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@PACZKA_ID_IN", paczkaId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_KOMPLETACJA_PACZKA_ZAMKNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaPaczkaRozpakuj(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal paczkaId As Integer) As KompletacjaPaczkiWynik

        Dim wynik As New KompletacjaPaczkiWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim iloscOut As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@PACZKA_ID_IN", paczkaId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_KOMPLETACJA_PACZKA_ROZPAKUJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KompletacjaZakoncz(ByVal sesja As Byte(), ByVal zamowienieId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim iloscOut As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_KOMPLETACJA_ZAKONCZ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Dokumenty"

    <WebMethod()> _
    Public Function DokumentyDoWydrukuListaPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_DO_WYGENEROWANIA_LISTA_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentyDodatkoweDoWydrukuListaPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_INNE_DO_WYGENEROWANIA_LISTA_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function





    <WebMethod()> _
    Public Function DokumentyAtrybutyWydrukuListaPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_ATRYBUTY_WYDRUKU_LISTA_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentyDoWydrukuDanePobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer,
                                      ByVal dokumentTypId As Integer,
                                      ByVal dokumentTypGrupaId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokumentId As Integer = 0
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
                htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)
                htParams.Add("@DOKUMENT_DATA", Now())
                htParams.Add("@OUT_DOKUMENT_ID", dokumentId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_GENERUJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentZewnetrznyPobierz(ByVal sesja As Byte(),
                                              ByVal projektID As Integer,
                                              ByVal zamowienieId As Integer,
                                              ByVal dataWystawienia As DateTime,
                                              ByVal dokumentTypId As Integer,
                                              ByVal dokumentTypGrupaId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokumentId As Integer = 0
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID", projektID)
                htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
                htParams.Add("@DOKUMENT_DATA", IIf(dataWystawienia = DateTime.MinValue, DBNull.Value, dataWystawienia))
                htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)

                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_ZEW_GENERUJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentInnyDanePobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer,
                                      ByVal dokumentTypId As Integer,
                                      ByVal dokumentTypGrupaId As Integer,
                                      ByVal dokumentId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokumentOutId As Integer = 0
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
                htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)
                htParams.Add("@ID", dokumentId)
                htParams.Add("@OUT_DOKUMENT_ID", dokumentOutId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_INNE_GENERUJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function DokumentWZNumerPobierz(ByVal sesja As Byte(),
                                           ByVal zamowienieId As Integer) As DokumentWZNumerWynik

        Dim wynik As New DokumentWZNumerWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokWZNumer As String = ""

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
                htParams.Add("@OUT_DOC_WZ_NR:100", dokWZNumer)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_WZ_NR_POBIERZ", htParams)
                'dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                If procRes.Status = 0 Then
                    wynik.dokumentWZNumer = NZ(htParams("@OUT_DOC_WZ_NR:100"), "")
                End If
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function KurierDokumentDodajZmien(ByVal sesja As Byte(),
                                             ByVal zamowienieId As Integer,
                                             ByVal dokumentTypId As Integer,
                                             ByVal dokumentTypGrupaId As Integer,
                                             ByVal dokumentNr As String) As StatusWynik

        Dim wynik As New StatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokWZNumer As String = ""

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@DOKUMENT_TYP_ID_IN", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID_IN", dokumentTypGrupaId)
                htParams.Add("@DOKUMENT_NR_IN", dokumentNr)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_KURIER_DODAJ_ZMIEN", htParams)
                'dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentPlikDodaj(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal dokumentTypId As Integer, ByVal plik As Byte(), ByVal dokumentTypGrupaId As Integer) As DokumentPlikDodajWynik
        Dim wynik As New DokumentPlikDodajWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@DOKUMENT_TYP_ID_IN", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID_IN", dokumentTypGrupaId)
                htParams.Add("@PLIK_IN", plik)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_PLIK_WSTAW", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DokumentPlikPobierz(ByVal sesja As Byte(), ByVal klientId As Integer, ByVal projektId As Integer, ByVal dokumentTypId As Integer, ByVal zamowienieNr As String, ByVal dokumentTypGrupaId As Integer) As DokumentPlikPobierzWynik

        Dim wynik As New DokumentPlikPobierzWynik
        Dim dSet As DataSet = New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@KLIENT_ID_IN", klientId)
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_NR_IN", zamowienieNr)
                htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_PLIK_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotDokumentyDoWydrukuListaPobierz(ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer, ByVal zwrotId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@ZWROT_ID_IN", zwrotId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_DOK_DO_WYGENEROWANIA_LISTA_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotKorektaPlikDodaj(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal ZwrotId As Integer, ByVal dokumentTypId As Integer, ByVal dokumentTypGrupaId As Integer, ByVal plik As Byte()) As DokumentPlikDodajWynik
        Dim wynik As New DokumentPlikDodajWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
                htParams.Add("@ZWROT_ID_IN", ZwrotId)
                htParams.Add("@DOKUMENT_TYP_ID_IN", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID_IN", dokumentTypGrupaId)
                htParams.Add("@PLIK_IN", plik)
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_FKOREKTA_PLIK_WSTAW", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DokumentDanePrzyjeciaWydaniaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dokumentNr As String, ByVal dokumentZnacznik As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projekt_id)
                htParams.Add("@DOKUMENT_NR_IN", dokumentNr)
                htParams.Add("@DOKUMENT_ZNACZNIK_IN", dokumentZnacznik)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_Q_DANE_PRZYJECIE_WYDANIE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function DokumentSzablonWydrukuDodaj(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dokumentGrupaID As String, ByVal dokumentNazwa As String, ByVal dokumentOpis As String, dokumentDane As String) As SzablonWydrukuPlikDodajWynik
        Dim wynik As New SzablonWydrukuPlikDodajWynik
        Using cnn As New ConnectionSuperPaker


            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DOKUMENT_SZABLON_GRUPA_ID_IN", dokumentGrupaID)
            htParams.Add("@DOK_SZABLON_NAZWA_IN", dokumentNazwa)
            htParams.Add("@DOK_SZABLON_OPIS_IN", dokumentOpis)
            htParams.Add("@DOKUMENT_SZABLON_DANE_IN", dokumentDane)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_SZABLON_DODAJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonWydrukuPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As SzablonWydrukuPobierzWynik
        Dim wynik As New SzablonWydrukuPobierzWynik

        Using cnn As New ConnectionSuperPaker

            Dim dSet As DataSet = New DataSet
            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_SZABLON_LISTA_POBIERZ", htParams, dSet)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DokumentInfoPobierz(ByVal sesja As Byte(), ByVal klientId As Integer, ByVal projektId As Integer, zamowienieNr As String, dokumentTypId As Integer, dokumentTypGrupaId As Integer) As DokumentInfoPobierzWynik
        Dim wynik As New DokumentInfoPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokumentDane As String = ""

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@KLIENT_ID_IN", klientId)
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@ZAMOWIENIE_NR_IN", zamowienieNr)
            htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
            htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)
            htParams.Add("@OUT_DOKUMENT_INFO:-1", dokumentDane)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_DOK_INFO_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.dokument_dane = NZ(htParams("@OUT_DOKUMENT_INFO:-1"), "")
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try
        End Using

        Return wynik
    End Function
#End Region

#Region "Produkty"

    <WebMethod()> _
    Public Function ProduktyListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ProduktyListaXMLPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As ProduktyListaXMLWynik

        Dim wynik As New ProduktyListaXMLWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim XML As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@OUT_PRODUKT_LISTA_XML:-1", XML)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_PRODUKT_LISTA_XML", htParams, dSet)
                wynik.dane = dSet
                wynik.xml = htParams("@OUT_PRODUKT_LISTA_XML:-1")
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktyTypPobierz(ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_TYP_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function StawkiVatPobierz(ByVal sesja As Byte()) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_STAWKI_VAT_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktPrefixPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As ProduktPrefixPobierzWynik

        Dim wynik As New ProduktPrefixPobierzWynik



        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            Dim projektPrefixOUT As String = ""
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""


            Dim htParams As Hashtable = New Hashtable()

            htParams.Add("@OUT_PREFIX_ID:25", projektPrefixOUT)
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PRODUKT_PREFIX_POBIERZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.prefix = NZ(htParams("@OUT_PREFIX_ID:25"), "")
                End If


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktDodaj(ByVal sesja As Byte(), ByVal dane As DataSet, ByVal projekt_id As Integer, ByVal klasa_abc As String) As ProduktDodajWynik

        Dim wynik As New ProduktDodajWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim ProduktTable As New DataTable
            Dim produktIdOUT As Integer = 0

            If dane.Tables.Count > 0 Then
                ProduktTable = dane.Tables(0).Copy
            End If



            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@TableProdukt_in", ProduktTable)
                htParams.Add("@PROJEKT_ID_IN", projekt_id)
                htParams.Add("@KLASA_ABC_IN", klasa_abc)
                htParams.Add("@OUT_PRODUKT_ID", produktIdOUT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PRODUKT_DODAJ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.produktId = NZ(htParams("@OUT_PRODUKT_ID"), 0)
                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ProduktEdycjaZapisz(ByVal sesja As Byte(), ByVal dane As DataSet, ByVal projekt_id As Integer, ByVal klasa_abc As String, ByVal domyslny_dostawca_id As Integer) As ProduktEdycjaZapiszWynik

        Dim wynik As New ProduktEdycjaZapiszWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim ProduktTable As New DataTable
            Dim DostawcyTable As New DataTable
            Dim AtrybutyTable As New DataTable

            Dim produktIdOUT As Integer = 0

            If dane.Tables.Count > 0 Then
                ProduktTable = dane.Tables(0).Copy
            End If

            If dane.Tables.Count > 1 Then
                DostawcyTable = dane.Tables(1).Copy
            End If

            If dane.Tables.Count > 2 Then
                AtrybutyTable = dane.Tables(2).Copy
            Else
                AtrybutyTable.Columns.Add("PRODUKT_ID", GetType(Integer)).SetOrdinal(0)
                AtrybutyTable.Columns.Add("ATRYBUT_ID", GetType(Integer)).SetOrdinal(1)
                AtrybutyTable.Columns.Add("NAZWA", GetType(String)).SetOrdinal(2)
                AtrybutyTable.Columns.Add("WARTOSC_NVARCHAR", GetType(String)).SetOrdinal(3)
                AtrybutyTable.Columns.Add("WARTOSC_BIN", GetType(Byte())).SetOrdinal(4)
            End If

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@TableProdukt_in", ProduktTable)
                htParams.Add("@TableDostawcy_in", DostawcyTable)
                htParams.Add("@TableAtrybutyProduktu_in", AtrybutyTable)
                htParams.Add("@PROJEKT_ID_IN", projekt_id)
                htParams.Add("@DOSTAWCA_ID_DOMYSLNY_IN", domyslny_dostawca_id)
                htParams.Add("@KLASA_ABC_IN", klasa_abc)
                htParams.Add("@OUT_PRODUKT_ID", produktIdOUT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PRODUKT_EDYCJA_ZAPISZ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.produktId = NZ(htParams("@OUT_PRODUKT_ID"), 0)
                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ProduktEdytuj(ByVal sesja As Byte(), ByVal ProduktEdytowanyId As Integer, projekt_id As Integer) As ProduktEdytujWynik

        Dim wynik As New ProduktEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PRODUKT_ID_EDYTOWANY_IN", ProduktEdytowanyId)
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_PRODUKT_EDYCJA_ROZPOCZNIJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If

            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktDaneLogistyczneDodaj(ByVal sesja As Byte(), ByVal dane As DataSet, ByVal PRODUKT_TYP_NAZWA_IN As String _
                                                , ByVal PROJEKT_ID_IN As Integer _
                                                , ByVal DOSTAWCA_DOMYSLNY_IN As Integer _
                                                , ByVal STAWKA_VAT_PROCENT_IN As Integer _
                                                , ByVal WAGA_NETTO_LV0_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV0_IN As Decimal _
                                                , ByVal DLUGOSC_LV0_IN As Decimal _
                                                , ByVal SZEROKOSC_LV0_IN As Decimal _
                                                , ByVal WYSOKOSC_LV0_IN As Decimal _
                                                , ByVal ILOSC_LV0_IN As Integer _
                                                , ByVal WAGA_NETTO_LV1_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV1_IN As Decimal _
                                                , ByVal DLUGOSC_LV1_IN As Decimal _
                                                , ByVal SZEROKOSC_LV1_IN As Decimal _
                                                , ByVal WYSOKOSC_LV1_IN As Decimal _
                                                , ByVal ILOSC_LV1_IN As Integer _
                                                , ByVal WAGA_NETTO_LV2_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV2_IN As Decimal _
                                                , ByVal DLUGOSC_LV2_IN As Decimal _
                                                , ByVal SZEROKOSC_LV2_IN As Decimal _
                                                , ByVal WYSOKOSC_LV2_IN As Decimal _
                                                , ByVal ILOSC_LV2_IN As Integer _
                                                , ByVal WAGA_NETTO_LV3_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV3_IN As Decimal _
                                                , ByVal DLUGOSC_LV3_IN As Decimal _
                                                , ByVal SZEROKOSC_LV3_IN As Decimal _
                                                , ByVal WYSOKOSC_LV3_IN As Decimal _
                                                , ByVal ILOSC_LV3_IN As Integer _
                                                , ByVal KLASA_ABC_IN As String) As ProduktEdycjaZapiszWynik


        Return ProduktDaneLogistyczneDodajV2(sesja, dane, PRODUKT_TYP_NAZWA_IN, PROJEKT_ID_IN, DOSTAWCA_DOMYSLNY_IN, STAWKA_VAT_PROCENT_IN, WAGA_NETTO_LV0_IN, WAGA_BRUTTO_LV0_IN, DLUGOSC_LV0_IN, SZEROKOSC_LV0_IN, WYSOKOSC_LV0_IN, ILOSC_LV0_IN, WAGA_NETTO_LV1_IN, WAGA_BRUTTO_LV1_IN, DLUGOSC_LV1_IN, SZEROKOSC_LV1_IN, WYSOKOSC_LV1_IN, ILOSC_LV1_IN, WAGA_NETTO_LV2_IN, WAGA_BRUTTO_LV2_IN, DLUGOSC_LV2_IN, SZEROKOSC_LV2_IN, WYSOKOSC_LV2_IN, ILOSC_LV2_IN, WAGA_NETTO_LV3_IN, WAGA_BRUTTO_LV3_IN, DLUGOSC_LV3_IN, SZEROKOSC_LV3_IN, WYSOKOSC_LV3_IN, ILOSC_LV3_IN, KLASA_ABC_IN, False, False)
    End Function

    <WebMethod()> _
    Public Function ProduktyDaneLogistycznePobierzMulti(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dane As String, ByVal opakowanie As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            If opakowanie.Length = 0 Then
                opakowanie = "SZTUKA"
            End If
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@XML_DANE_IN", dane)
            htParams.Add("@OPAKOWANIE_IN", opakowanie)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_DANE_LOGISTYCZNE_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktyDaneLogistycznePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal produkt_nr As String, ByVal opakowanie As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""
            Dim xml_dane As String = ""
            xml_dane = "<ROW PRODUKT_NR = """ & produkt_nr & """/>"

            If opakowanie.Length = 0 Then
                opakowanie = "SZTUKA"
            End If
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@XML_DANE_IN", xml_dane)
            htParams.Add("@OPAKOWANIE_IN", opakowanie)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_DANE_LOGISTYCZNE_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktListaDanePoSKUPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal produkty As DataSet) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableProdukty As DataTable = New DataTable

        If produkty.Tables.Count > 0 Then
            dTableProdukty = produkty.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@TableSku_in", dTableProdukty)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKT_POBIERZ_DANE_PO_SKU_LISTA", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktyStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal dataWprowadzeniaOd As DateTime,
                                      ByVal dataWprowadzeniaDo As DateTime,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer) As ProduktyStronaWynik

        Dim wynik As New ProduktyStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0
        Dim dTableStatusy As DataTable = New DataTable

       
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try


                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)

                htParams.Add("@DATA_OD", IIf(dataWprowadzeniaOd = DateTime.MinValue, DBNull.Value, dataWprowadzeniaOd))
                htParams.Add("@DATA_DO", IIf(dataWprowadzeniaDo = DateTime.MinValue, DBNull.Value, dataWprowadzeniaDo))

                    htParams.Add("@STRONA_NR_IN", stronaNumer)
                    htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                    'htParams.Add("@SORT_PO_IN", sortPo)
                    'htParams.Add("@SORT_KIERUNEK_IN", sortAsc)

                    htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                    'htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_PRODUKTY_LISTA_POBIERZ_STRONA", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    If procRes.Status = 0 Then
                        wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                    End If
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
    
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktDaneLogistyczneDodajV2(ByVal sesja As Byte(), ByVal dane As DataSet, ByVal PRODUKT_TYP_NAZWA_IN As String _
                                                , ByVal PROJEKT_ID_IN As Integer _
                                                , ByVal DOSTAWCA_DOMYSLNY_IN As Integer _
                                                , ByVal STAWKA_VAT_PROCENT_IN As Integer _
                                                , ByVal WAGA_NETTO_LV0_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV0_IN As Decimal _
                                                , ByVal DLUGOSC_LV0_IN As Decimal _
                                                , ByVal SZEROKOSC_LV0_IN As Decimal _
                                                , ByVal WYSOKOSC_LV0_IN As Decimal _
                                                , ByVal ILOSC_LV0_IN As Integer _
                                                , ByVal WAGA_NETTO_LV1_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV1_IN As Decimal _
                                                , ByVal DLUGOSC_LV1_IN As Decimal _
                                                , ByVal SZEROKOSC_LV1_IN As Decimal _
                                                , ByVal WYSOKOSC_LV1_IN As Decimal _
                                                , ByVal ILOSC_LV1_IN As Integer _
                                                , ByVal WAGA_NETTO_LV2_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV2_IN As Decimal _
                                                , ByVal DLUGOSC_LV2_IN As Decimal _
                                                , ByVal SZEROKOSC_LV2_IN As Decimal _
                                                , ByVal WYSOKOSC_LV2_IN As Decimal _
                                                , ByVal ILOSC_LV2_IN As Integer _
                                                , ByVal WAGA_NETTO_LV3_IN As Decimal _
                                                , ByVal WAGA_BRUTTO_LV3_IN As Decimal _
                                                , ByVal DLUGOSC_LV3_IN As Decimal _
                                                , ByVal SZEROKOSC_LV3_IN As Decimal _
                                                , ByVal WYSOKOSC_LV3_IN As Decimal _
                                                , ByVal ILOSC_LV3_IN As Integer _
                                                , ByVal KLASA_ABC_IN As String _
                        , ByVal WYMAGANIE_DATA_WAZNOSCI As Boolean _
                        , ByVal WYMAGANIE_PARTII As Boolean) As ProduktEdycjaZapiszWynik

        Dim wynik As New ProduktEdycjaZapiszWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim ProduktTable As New DataTable
            Dim DostawcyTable As New DataTable
            Dim AtrybutyTable As New DataTable

            Dim produktIdOUT As Integer = 0

            If dane.Tables.Count > 0 Then
                ProduktTable = dane.Tables(0).Copy
            End If

            If dane.Tables.Count > 1 Then
                DostawcyTable = dane.Tables(1).Copy
            End If

            If dane.Tables.Count > 2 Then
                AtrybutyTable = dane.Tables(2).Copy
            Else
                AtrybutyTable.Columns.Add("PRODUKT_ID", GetType(Integer)).SetOrdinal(0)
                AtrybutyTable.Columns.Add("ATRYBUT_ID", GetType(Integer)).SetOrdinal(1)
                AtrybutyTable.Columns.Add("NAZWA", GetType(String)).SetOrdinal(2)
                AtrybutyTable.Columns.Add("WARTOSC_NVARCHAR", GetType(String)).SetOrdinal(3)
                AtrybutyTable.Columns.Add("WARTOSC_BIN", GetType(Byte())).SetOrdinal(4)
            End If


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()


                htParams.Add("@TableProdukt_in", ProduktTable)
                htParams.Add("@TableDostawcy_in", DostawcyTable)
                htParams.Add("@TableAtrybutyProduktu_in", AtrybutyTable)
                htParams.Add("@PRODUKT_TYP_NAZWA_IN", PRODUKT_TYP_NAZWA_IN)
                htParams.Add("@PROJEKT_ID_IN", PROJEKT_ID_IN)
                htParams.Add("@DOSTAWCA_DOMYSLNY_IN", DOSTAWCA_DOMYSLNY_IN)
                htParams.Add("@STAWKA_VAT_PROCENT_IN", STAWKA_VAT_PROCENT_IN)
                htParams.Add("@WAGA_NETTO_LV0_IN", WAGA_NETTO_LV0_IN)
                htParams.Add("@WAGA_BRUTTO_LV0_IN", WAGA_BRUTTO_LV0_IN)
                htParams.Add("@DLUGOSC_LV0_IN", DLUGOSC_LV0_IN)
                htParams.Add("@SZEROKOSC_LV0_IN", SZEROKOSC_LV0_IN)
                htParams.Add("@WYSOKOSC_LV0_IN", WYSOKOSC_LV0_IN)
                htParams.Add("@ILOSC_LV0_IN", ILOSC_LV0_IN)
                htParams.Add("@WAGA_NETTO_LV1_IN", WAGA_NETTO_LV1_IN)
                htParams.Add("@WAGA_BRUTTO_LV1_IN", WAGA_BRUTTO_LV1_IN)
                htParams.Add("@DLUGOSC_LV1_IN", DLUGOSC_LV1_IN)
                htParams.Add("@SZEROKOSC_LV1_IN", SZEROKOSC_LV1_IN)
                htParams.Add("@WYSOKOSC_LV1_IN", WYSOKOSC_LV1_IN)
                htParams.Add("@ILOSC_LV1_IN", ILOSC_LV1_IN)
                htParams.Add("@WAGA_NETTO_LV2_IN", WAGA_NETTO_LV2_IN)
                htParams.Add("@WAGA_BRUTTO_LV2_IN", WAGA_BRUTTO_LV2_IN)
                htParams.Add("@DLUGOSC_LV2_IN", DLUGOSC_LV2_IN)
                htParams.Add("@SZEROKOSC_LV2_IN", SZEROKOSC_LV2_IN)
                htParams.Add("@WYSOKOSC_LV2_IN", WYSOKOSC_LV2_IN)
                htParams.Add("@ILOSC_LV2_IN", ILOSC_LV2_IN)
                htParams.Add("@WAGA_NETTO_LV3_IN", WAGA_NETTO_LV3_IN)
                htParams.Add("@WAGA_BRUTTO_LV3_IN", WAGA_BRUTTO_LV3_IN)
                htParams.Add("@DLUGOSC_LV3_IN", DLUGOSC_LV3_IN)
                htParams.Add("@SZEROKOSC_LV3_IN", SZEROKOSC_LV3_IN)
                htParams.Add("@WYSOKOSC_LV3_IN", WYSOKOSC_LV3_IN)
                htParams.Add("@ILOSC_LV3_IN", ILOSC_LV3_IN)

                htParams.Add("@KLASA_ABC_IN", KLASA_ABC_IN)
                htParams.Add("@OUT_PRODUKT_ID", produktIdOUT)

                htParams.Add("@WYM_DATA_WAZNOSCI", WYMAGANIE_DATA_WAZNOSCI)
                htParams.Add("@WYM_PARTIA", WYMAGANIE_PARTII)


                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PRODUKT_DANE_LOGISTYCZNE_DODAJ", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally

                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If procRes.Status = 0 Then
                        wynik.produktId = NZ(htParams("@OUT_PRODUKT_ID"), 0)
                    End If
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.status = -1
            wynik.status_opis = "Brak danych do przetworzenia"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ProduktyMasowaEdycjaPolaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As ProduktyMasowaEdycjaPolaListaPobierzWynik

        Dim wynik As New ProduktyMasowaEdycjaPolaListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_MASOWA_EDYCJA_POLA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktyMasowaEdycjaZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zmianaPoleId As Integer, ByVal warunekZ As String, ByVal warunekNa As String, ByVal produkty As DataSet) As ProduktyMasowaEdycjaZapiszWynik

        Dim wynik As New ProduktyMasowaEdycjaZapiszWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableProdukty As DataTable = New DataTable

        If produkty.Tables.Count > 0 Then
            dTableProdukty = produkty.Tables(0).Copy
        End If

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@PRODUKT_MASOWE_ZMIANY_ID", zmianaPoleId)
            htParams.Add("@WARUNEK_Z", warunekZ)
            htParams.Add("@WARUNEK_NA", warunekNa)
            htParams.Add("@TableProdukt_in", dTableProdukty)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKTY_MASOWA_EDYCJA_ZAPISZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
#End Region

#Region "ABCDATA"
    <WebMethod()> _
    Public Function ProduktyAbcDataListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKT_ABCDATA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ProduktyAbcDataListaZapisz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dane As String) As StatusWynik

        Dim wynik As New StatusWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DANE_IN", dane)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PRODUKT_ABCDATA_LISTA_ZAPISZ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowieniaAbcdataDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As ZamowieniaAbcdataDanePobierzWynik

        Dim wynik As New ZamowieniaAbcdataDanePobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim agregatIdOUT As Integer = 0

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@OUT_AGREGAT_ID", agregatIdOUT)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_ZAMOWIENIE_ABCDATA_DANE_POBIERZ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.agregat_id = NZ(htParams("@OUT_AGREGAT_ID"), 0)
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowieniaAbcdataBlad(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal agregat_id As Integer, ByVal komunikat As String) As StatusWynik

        Dim wynik As New StatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim agregatIdOUT As Integer = 0

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@AGREGAT_ID_IN", agregat_id)
            htParams.Add("@KOMUNIKAT_IN", komunikat)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIA_ABCDATA_BLAD", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowieniaAbcdataZlozone(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal agregat_id As Integer, ByVal numer_zamowienia As String) As StatusWynik

        Dim wynik As New StatusWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim agregatIdOUT As Integer = 0

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@AGREGAT_ID_IN", agregat_id)
            htParams.Add("@NUMER_ZAMOWIENIA_IN", numer_zamowienia)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIA_ABCDATA_ZLOZONE", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
#End Region

#Region "Raporty"

    <WebMethod()> _
    Public Function RaportSprzedazyDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_SPRZEDAZ_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportSprzedazyPartiaDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_SPRZEDAZ_PLUS_PARTIA_DATA_WAZNOSCI_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportSprzedazProduktDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_SPRZEDAZ_PRODUKT_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZalacznikVATDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ZALACZNIK_VAT_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZalacznikVATZbiorczyDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ZALACZNIK_VAT_ZBIORCZY_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportStanyMagazynoweDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_STANY_MAGAZYNOWE_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
 

    <WebMethod()> _
    Public Function RaportStanyMagazynoweBiezaceDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_STANY_MAGAZYNOWE_BIEZACE_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportStanyMagazynoweBiezaceDataWaznosciPartiaDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_STANY_MAGAZYNOWE_BIEZACE_PLUS_DATA_WAZNOSCI_PARTIA_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportRuchyMiedzymagazynowePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime, ByVal statusyDS As DataSet) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable
        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projekt_id)
                htParams.Add("@DATA_OD_IN", dataOd)
                htParams.Add("@DATA_DO_IN", dataDo)
                htParams.Add("@STATUSY_LISTA_IN", dTableStatusy) 'FIXIT - POPRAWIĆ NA TAKIE JAK W POPRAWIONEJ PROCEDURZE
                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RUCHY_MIEDZYMAGAZYNOWE_GENERUJ", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message


                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane do przetworzenia"
        End If

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportWejsciaWyjsciaDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_WEJSCIA_WYJSCIA_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZwrotowDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ZWROT_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportFakturowniaPLDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_FAKTUROWNIA_PL_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportDarmowaDostawaDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_DARMOWA_DOSTAWA_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportFakturXMLDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportXMLWynik

        Dim wynik As New RaportXMLWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim XML As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@OUT_XML:-1", XML)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_RPT_FAKTUR_XML_GENERUJ", htParams, dSet)
                wynik.dane = dSet
                wynik.xml = htParams("@OUT_XML:-1")
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportJpkGeneruj(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportXMLWynik

        Dim wynik As New RaportXMLWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim xmlDoc As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@OUT_XML:-1", xmlDoc)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_GENERUJ", htParams, dSet)
                wynik.dane = dSet
                wynik.xml = htParams("@OUT_XML:-1")
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportJpkGenerujNiesklasyfikowane(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_GENERUJ_NIESKLASYFIKOWANE", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportJpkXmlWysylka(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim xmlDoc As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@OUT_XML:-1", xmlDoc)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_XML_DOWYSYLKI", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportJpkNieklasyfikowaneEdytuj(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal doc_Id As Integer, ByVal doc_Type As String, daneTablica As DataSet _
                                               , skad As String, dokad As String, numerFa As String, dataFa As String) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DOC_ID", doc_Id)
            htParams.Add("@DOC_TYPE", doc_Type)
            htParams.Add("@TABLE_ITEMS", daneTablica.Tables(0))
            htParams.Add("@SKAD", skad)
            htParams.Add("@DOKAD", dokad)
            htParams.Add("@NUMER_FA", numerFa)
            htParams.Add("@DATA_FA", dataFa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_UPDATE_NIEKLASYFIKOWANE", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportJpkNieklasyfikowaneEdytujWrocePozniej(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal doc_Id As Integer, daneTablica As DataSet) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DOC_ID", doc_Id)
            htParams.Add("@TABLE_ITEMS", daneTablica.Tables(0))

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_UPDATE_NIEKLASYFIKOWANE_WROCE_POZNIEJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportJpkNieklasyfikowaneEdytujNieDoJPK(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal doc_Id As Integer) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DOC_ID", doc_Id)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_UPDATE_NIEKLASYFIKOWANE_NIE_DO_JPK", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportJpkZapiszWygenerowanyPlik(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime, ByVal XmlJpkDoc As String) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim xmlDoc As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("XML_JPK_DOC", XmlJpkDoc)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_ZAPISZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportJpkUpdateWysylka(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal xml_Id As Integer, ByVal saveJpkExecution As SaveJpkExecution) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim xmlDoc As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DOC_ID", xml_Id)
            htParams.Add("@AZURE_CODE", saveJpkExecution.upoResponse.Code)
            htParams.Add("@AZURE_DESCRIPTION", saveJpkExecution.upoResponse.Description)
            htParams.Add("@AZURE_DETAILS", saveJpkExecution.upoResponse.Details)
            htParams.Add("@AZURE_TIMESTAMP", saveJpkExecution.upoResponse.Timestamp)
            htParams.Add("@AZURE_UPO", saveJpkExecution.upoResponse.Upo)
            htParams.Add("@XML_WYSLANY", saveJpkExecution.fileXml)
            htParams.Add("@AZURE_REF_NUM", saveJpkExecution.refNum)


            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_UPDATE_WYSYLKA", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportJpkUpdateRefnum(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal doc_Id As Integer, ByVal refNum As String) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DOC_ID", doc_Id)
            htParams.Add("@AZURE_REF_NUM", refNum)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_JPK_UPDATE_REFNUM", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportKosztDostawyDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_KOSZT_DOSTAWY_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportKpwKpiXmlGeneruj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal data As DateTime) As RaportXMLWynik

        Dim wynik As New RaportXMLWynik


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim XML As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_IN", data)
            htParams.Add("@OUT_XML:-1", XML)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_RPT_KPW_KPI_XML_GENERUJ", htParams)
                wynik.xml = htParams("@OUT_XML:-1")
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function


    <WebMethod()> _
    Public Function RaportTerminowoscDostawDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_TERMINOWOSC_DOSTAW_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function RaportTerminowoscZlecenDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_TERMINOWOSC_ZLECEN_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function RaportAnulowanieZlecenDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ANULOWANIE_ZLECEN_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function RaportRozliczeniaDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ROZLICZENIE_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function RaportBrakiWZamowieniachDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_BRAKI_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportWejsciaWyjsciaProduktDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dataOd As DateTime, ByVal dataDo As DateTime, ByVal produktNr As String, ByVal magazynyWirtualne As DataSet) As RaportWynik
        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim magazynyWirtualneDT As DataTable

        Using cnn As New ConnectionSuperPaker

            If Not magazynyWirtualne Is Nothing AndAlso magazynyWirtualne.Tables.Count > 0 Then
                magazynyWirtualneDT = magazynyWirtualne.Tables(0).Copy
            Else
                wynik.status = -1
                wynik.status_opis = "Nie przesłano listy magazynów wirtualnych"
                Return wynik
            End If

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@PRODUKT_NR_IN", produktNr)
            htParams.Add("@MAG_WIRT_LISTA_IN", magazynyWirtualneDT)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_WEJSCIA_WYJSCIA_PRODUKT_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZamowienXMLDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal orderIdFrom As Integer) As RaportXMLWynik

        Dim wynik As New RaportXMLWynik
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult() With {.Status = DataAccess.Status.Error, .Message = ""}
            Dim XML As String = ""
            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ORDER_ID_FROM", orderIdFrom)
            htParams.Add("@OUT_XML:-1", XML)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDSOutput(cnn, sesja, "UP_ZAMOWIENIE_KAMSOFT_WYSLIJ_XML", htParams, dSet)
                wynik.dane = dSet
                wynik.xml = htParams("@OUT_XML:-1")
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportFakturKorygujacychDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RAPORT_FAKTUR_KORYGUJACYCH", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function RaportWmsDziennaLiczbaZamowienDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_DZIENNA_LICZBA_ZAMOWIEN", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportStatystykiDostawDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RAPORT_STATYSTYKI_DOSTAW_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportParagonowDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RAPORT_PARAGONOW_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZamowieniaPozycjeDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ZAMOWIENIA_POZYCJE_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportSpakowanychZamowienDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime, ByVal filtry As DataSet) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim projektyTbl As DataTable

        If Not filtry Is Nothing AndAlso filtry.Tables.Count > 0 Then
            projektyTbl = filtry.Tables(0).Copy
        Else
            wynik.status = -1
            wynik.status_opis = "Nie przesłano listy filtrów"
            Return wynik
        End If


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@PROJEKTY", projektyTbl)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ALL_SPAKOWANE_ZAMOWIENIA_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportZamowieniaPozycjePartiaWaznosciDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_ZAMOWIENIA_POZYCJE_PARTIA_DATA_WAZNOSCI_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportFakturyParagonyPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer,
                                               ByVal dataOd As DateTime, ByVal dataDo As DateTime) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_FAKTURY_PARAGONY_XML_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportDowolnyListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal raportTypId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@RAPORT_TYP_ID", raportTypId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RAPORTY_DLA_PROJEKTU_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportDowolnyPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal raportId As Integer, ByVal raportTypId As Integer,
                                               ByVal filtrXml As String) As RaportWynik

        Dim wynik As New RaportWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@RAPORT_ID", raportId)
            htParams.Add("@RAPORT_TYP_ID", raportTypId)
            htParams.Add("@FILTR_XML", filtrXml)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RPT_RAPORT_DLA_PROJEKTU_DOWOLNY", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RaportDowolnyFiltrListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal raportId As Integer, ByVal raportTypId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@RAPORT_ID", raportId)
            htParams.Add("@RAPORT_TYP_ID", raportTypId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_RAPORT_DOWOLNY_FILTR_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function



#End Region

#Region "Rozliczenia"
    <WebMethod()> _
    Public Function RozliczeniaNiezamknieteDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As RozliczeniaWynik

        Dim wynik As New RozliczeniaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_DPD_PACZKA_NIEZAMKNIETA_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function RozliczeniaZamknij(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal kwota As String, ByVal data_przelewu As String, ByVal data_zamkniecia As String) As RozliczeniaZamknijWynik
        Dim wynik As New RozliczeniaZamknijWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@KWOTA_IN", kwota)
            htParams.Add("@DATA_PRZELEWU_IN", data_przelewu)
            htParams.Add("@DATA_ZAMKNIECIA_IN", data_zamkniecia)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_TRN_DPD_PACZKA_ZAMKNIJ", htParams)
                'wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODDodaj(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal plikId As Integer, ByVal plikCODTypId As Integer, ByVal plikNazwa As String, ByVal plikDane As DataSet) As RozliczeniePlikCODDodajWynik
        Dim wynik As New RozliczeniePlikCODDodajWynik
        Dim dSet As DataSet = New DataSet
        Dim plikPozycjeDT As DataTable
        If plikDane.Tables.Count > 0 Then
            plikPozycjeDT = plikDane.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID", projektID)
                htParams.Add("@TRN_PLIK_COD_RODZAJ_ID", plikCODTypId)
                htParams.Add("@PLIK_NAZWA", plikNazwa)
                htParams.Add("@TablePlikCOD", plikPozycjeDT)
                htParams.Add("@TRN_PLIK_COD_ID", plikId)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_DODAJ", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane do zapisania"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODRodzajListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As RozliczeniePlikCODRodzajListaPobierzWynik
        Dim wynik As New RozliczeniePlikCODRodzajListaPobierzWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_RODZAJ_LISTA", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODStatusListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As RozliczeniePlikCODStatusListaPobierzWynik
        Dim wynik As New RozliczeniePlikCODStatusListaPobierzWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_STATUS_LISTA", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODListaPlikowPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dataOd As DateTime) As RozliczeniePlikCODListaPlikowPobierzWynik
        Dim wynik As New RozliczeniePlikCODListaPlikowPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@DATA_OD", dataOd)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_LISTA", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODPozycjaStanListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As RozliczeniePlikCODPozycjaStanListaPobierzWynik
        Dim wynik As New RozliczeniePlikCODPozycjaStanListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_POZYCJA_STAN_LISTA", htParams, dSet)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODWalidacjaSlownikListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As RozliczeniePlikCODWalidacjaSlownikListaPobierzWynik
        Dim wynik As New RozliczeniePlikCODWalidacjaSlownikListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_POZYCJA_WALIDACJA_LISTA", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODPlikWalidacjaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikCODId As Integer) As RozliczeniePlikCODPlikWalidacjaPobierzWynik
        Dim wynik As New RozliczeniePlikCODPlikWalidacjaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@TRN_PLIK_COD_ID", plikCODId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_WALIDUJ", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODPlikSzczegolyPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikCODId As Integer) As RozliczeniePlikCODPlikSzczegolyPobierzWynik
        Dim wynik As New RozliczeniePlikCODPlikSzczegolyPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@TRN_PLIK_COD_ID", plikCODId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_SZCZEGOLY", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODOdrzucPlik(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikCODId As Integer) As RozliczeniePlikCODOdrzucPlikWynik
        Dim wynik As New RozliczeniePlikCODOdrzucPlikWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@TRN_PLIK_COD_ID", plikCODId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_ODRZUC ", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePlikCODDecyzjeZapisz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal plikId As Integer, ByVal decyzjeDS As DataSet) As RozliczeniePlikCODDecyzjeZapiszWynik
        Dim wynik As New RozliczeniePlikCODDecyzjeZapiszWynik
        Dim dSet As DataSet = New DataSet
        Dim plikPozycjeDT As DataTable
        If decyzjeDS.Tables.Count > 0 Then
            plikPozycjeDT = decyzjeDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID", projektID)
                htParams.Add("@TRN_PLIK_COD_ID", plikId)
                htParams.Add("@POZYCJA_DECYZJA", plikPozycjeDT)


                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_COD_POZYCJA_AKCEPTUJ", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane do zapisania"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePrzelewyMozliwePozycjePobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As RozliczeniePrzelewyMozliwePozycjePobierzWynik
        Dim wynik As New RozliczeniePrzelewyMozliwePozycjePobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_POZYCJE_DO_PRZELEWU", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePrzelewPlikDodajZmien(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal plikId As Integer, ByVal zatwierdz As Boolean, ByVal plikPozycjeDS As DataSet) As RozliczeniePrzelewPlikDodajZmienWynikWynik
        Dim wynik As New RozliczeniePrzelewPlikDodajZmienWynikWynik
        Dim dSet As DataSet = New DataSet
        Dim plikPozycjeDT As DataTable
        If plikPozycjeDS.Tables.Count > 0 Then
            plikPozycjeDT = plikPozycjeDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID", projektID)
                htParams.Add("@TRN_PLIK_PRZELEW_ID", plikId)
                htParams.Add("@POZYCJA_DO_PRZELEWU", plikPozycjeDT)
                htParams.Add("@ZATWIERDZ", zatwierdz)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_PRZELEW_UTWORZ_PLIK", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane do zapisania"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePrzelewPlikZatwierdz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal plikId As Integer, ByVal plikNazwa As String) As RozliczeniePrzelewPlikZatwierdzWynik
        Dim wynik As New RozliczeniePrzelewPlikZatwierdzWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektID)
            htParams.Add("@TRN_PLIK_PRZELEW_ID", plikId)
            htParams.Add("@PLIK_NAZWA", plikNazwa)


            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_PRZELEW_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePrzelewListaPlikowOczekujacychPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dataOd As DateTime) As RozliczeniePrzelewListaPlikowOczekujacychPobierzWynik
        Dim wynik As New RozliczeniePrzelewListaPlikowOczekujacychPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@DATA_OD", dataOd)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIKI_DO_PRZELEWU", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function RozliczeniePrzelewPlikOczekujacySzczegolyPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikCODId As Integer) As RozliczeniePrzelewPlikOczekujacySzczegolyPobierzWynik
        Dim wynik As New RozliczeniePrzelewPlikOczekujacySzczegolyPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@TRN_PLIK_PRZELEW_ID", plikCODId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_TRN_PLIK_POZYCJE_PLIKU_DO_PRZELEWU", htParams, dSet)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function



#End Region

#Region "Schemat"

    <WebMethod()> _
    Public Function SchemaatAtrybutListaPobierz(ByVal sesja As Byte(), ByVal zamowienie_id As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienie_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SCHEMAT_ATRYBUT_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function SchematZmianaMozliweListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zmienianySchematId As Integer, ByVal zamowienieId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@SCHEMAT_ID", zmienianySchematId)
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SCHEMAT_ZMIANA_LISTA", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function SchematZmianaZamowieniePojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal schematIdNowy As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            htParams.Add("@SCHEMAT_ID_NOWY", schematIdNowy)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_SCHEMAT_ZMIEN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function SchematZmianaZamowienieLista(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dane As DataSet, ByVal schematIdNowy As Integer) As StatusWynik

        Dim wynik As New StatusWynik

        If dane.Tables.Count > 0 AndAlso dane.Tables(0).Rows.Count > 0 Then

            Dim zamowienieTable As New DataTable

            zamowienieTable = dane.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID", projektId)
                htParams.Add("@ZAMOWIENIE_LISTA_IN", zamowienieTable)
                htParams.Add("@SCHEMAT_ID_NOWY", schematIdNowy)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_LISTA_SCHEMAT_ZMIEN", htParams)

                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowa lista zamówień do zmiany schematu"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieKurierZmianaMozliweListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zmienianySchematId As Integer, ByVal zamowienieId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@SCHEMAT_ID", zmienianySchematId)
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_KURIER_ZMIANA_LISTA", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZamowienieKurierZmianaZamowieniePojedynczo(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal kurierIdNowy As Integer, ByVal daneDodatkoweXml As String) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            htParams.Add("@KURIER_ID", kurierIdNowy)
            htParams.Add("@DANE_XML", daneDodatkoweXml)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_KURIER_ZMIEN ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

#End Region

#Region "Statystyki"

    <WebMethod()> _
    Public Function StatystykiProjektDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As StatystykiProjektWynik

        Dim wynik As New StatystykiProjektWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_STATS_PROJEKT_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function StatystykiWydajnoscDanePobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dataOd As DateTime, ByVal dataDo As DateTime) As StatystykiWydajnoscWynik

        Dim wynik As New StatystykiWydajnoscWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_STATS_WYDAJNOSC_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message


                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function StatystykiWydajnoscDaneStronaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal dataOd As DateTime, ByVal dataDo As DateTime,
                                                         ByVal stronaNumer As Integer,
                                                         ByVal stronaWielkosc As Integer,
                                                         ByVal sortPo As String,
                                                         ByVal sortAsc As Boolean) As StatystykiWydajnoscStronaWynik

        Dim wynik As New StatystykiWydajnoscStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@DATA_OD_IN", dataOd)
            htParams.Add("@DATA_DO_IN", dataDo)
            htParams.Add("@STRONA_NR_IN", stronaNumer)
            htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
            htParams.Add("@SORT_PO_IN", sortPo)
            htParams.Add("@SORT_KIERUNEK_IN", sortAsc)
            htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_STATS_WYDAJNOSC_DANE_STRONA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                If procRes.Status = 0 Then
                    wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                End If

                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
#End Region

#Region "Wydruk"
    <WebMethod()> _
    Public Function DrukarkaStawkaVATListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As DrukarkaStawkaVATListaWynik
        Dim wynik As New DrukarkaStawkaVATListaWynik
        Dim sData As New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_FISKALIZACJA_DRUKARKA_STAWKA_VAT_LISTA", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DrukarkaDaneFiskalneZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, klient_numer_zamowienia As String, drukarka_nr As String, drukarka_paragon_header As String, drukarka_paragon_nr As String, drukarka_paragon_vat_a As String, drukarka_paragon_vat_b As String, drukarka_paragon_vat_c As String, drukarka_paragon_vat_d As String) As DrukarkaDaneFiskalneZapiszWynik
        Dim wynik As New DrukarkaDaneFiskalneZapiszWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", klient_numer_zamowienia)
            htParams.Add("@DRUKARKA_NR_IN", drukarka_nr)
            htParams.Add("@DRUKARKA_PARAGON_HEADER_IN", drukarka_paragon_header)
            htParams.Add("@DRUKARKA_PARAGON_NR_IN", drukarka_paragon_nr)
            htParams.Add("@DRUKARKA_PARAGON_VAT_A_IN", drukarka_paragon_vat_a)
            htParams.Add("@DRUKARKA_PARAGON_VAT_B_IN", drukarka_paragon_vat_b)
            htParams.Add("@DRUKARKA_PARAGON_VAT_C_IN", drukarka_paragon_vat_c)
            htParams.Add("@DRUKARKA_PARAGON_VAT_D_IN", drukarka_paragon_vat_d)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_PARAGON_DANE_FISKALNE_ZAPISZ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SprawdzVatParagonu(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal numer_zamowienia As String, ByVal kwota_vat As Double) As SprawdzVatParagonuWynik

        Dim wynik As New SprawdzVatParagonuWynik
        Dim roznica As Double = 0D

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", numer_zamowienia)
            htParams.Add("@PARAGON_WARTOSC_VAT_IN", kwota_vat)
            htParams.Add("@OUT_WARTOSC_VAT_ROZNICA", roznica)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_SPRAWDZ_VAT_ZAMOWIENIA", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                ' wynik.roznica = IIf(IsDBNull(cmd.Parameters("roznica").Value), 999, cmd.Parameters("roznica").Value)
                If procRes.Status = 0 Then
                    wynik.roznica = NZ(htParams("@OUT_WARTOSC_VAT_ROZNICA"), 999)
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

#End Region

#Region "Ustawienia"
    <WebMethod()>
    Public Function ProjektUstawieniaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As ProjektUstawieniaWynik
        Dim wynik As New ProjektUstawieniaWynik
        Dim sData As New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_PROJEKT_USTAWIENIA_POBIERZ", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()>
    Public Function ProjektUstawieniaSprawdzDostepnoscKonfiguracji(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal uprawnienie As String) As StanowiskoUstawieniaPobierzWynik
        Dim wynik As New StanowiskoUstawieniaPobierzWynik
        Dim sData As New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@UPRAWNIENIE_IN", uprawnienie)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_USTAWIENIE_PROJEKT_SPRAWDZ_UPRAWNIENIE", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SchematKurierListaPobierz(ByVal sesja As Byte()) As ProjektUstawieniaWynik
        Dim wynik As New ProjektUstawieniaWynik
        Dim sData As New DataSet
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SCHEMAT_KURIER_LISTA_POBIERZ", htParams, sData)
                wynik.dane = sData
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function StanowiskoUstawieniaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal stanowiskoNazwa As String) As StanowiskoUstawieniaPobierzWynik
        Dim wynik As New StanowiskoUstawieniaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@NAZWA_STANOWISKA", stanowiskoNazwa)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_USTAWIENIE_STANOWISKA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function UstawieniaProjektPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer) As UstawieniaProjektPobierzWynik
        Dim wynik As New UstawieniaProjektPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_USTAWIENIE_PROJEKT_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function UstawieniaProjektZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal TypCFGId As Integer, ByVal wartoscPrzekazana As String, ByVal configTableId As Integer, ByVal wielokrotne As Boolean) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@CFG_ID_IN", TypCFGId)
            htParams.Add("@WARTOSC_IN", wartoscPrzekazana)
            htParams.Add("@CFG_VALUE_ID_IN", configTableId)
            htParams.Add("@WIELOKROTNE_IN", wielokrotne)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_USTAWIENIE_PROJEKT_ZAPISZ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function UstawieniaProjektUsun(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal configTableId As Integer) As StatusWynik

        Dim wynik As New StatusWynik
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@CFG_VALUE_ID_IN", configTableId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_USTAWIENIE_PROJEKT_USUN", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function StanowiskoUstawienieDodajEdytuj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal stanowskoNazwa As String, ByVal nazwaUstawienia As String, wartoscUstawienia As String) As StanowiskoUstawienieDodajEdytujWynik

        Dim wynik As New StanowiskoUstawienieDodajEdytujWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@NAZWA_STANOWISKA", stanowskoNazwa)
            htParams.Add("@NAZWA_USTAWIENIA", nazwaUstawienia)
            htParams.Add("@WARTOSC", wartoscUstawienia)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_USTAWIENIE_STANOWISKA_EDYTUJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function StanowiskoUstawienieUsun(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal stanowskoNazwa As String, ByVal nazwaUstawienia As String, wartoscUstawienia As String) As StanowiskoUstawienieUsunWynik

        Dim wynik As New StanowiskoUstawienieUsunWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""
            Dim magazynIdOUT As Integer = 0

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projekt_id)
            htParams.Add("@NAZWA_STANOWISKA", stanowskoNazwa)
            htParams.Add("@NAZWA_USTAWIENIA", nazwaUstawienia)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_USTAWIENIE_STANOWISKA_USUN", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

#End Region

#Region "Zdjęcia zamówień"
    Private Function PodajLiczbePlikowWKatalogu(ByVal katalog As String) As Integer
        StworzKatalog(katalog + "\")
        Dim iloscPlikow As Integer = 0
        If Directory.Exists(katalog) Then
            iloscPlikow = Directory.GetFiles(katalog + "\", "*.*").Length
        End If
        Return iloscPlikow
    End Function
    Private Function PodajPodkatalogINazweZdjecia(ByVal projektId As Integer, ByVal zamowienieId As Integer, ByVal extension As String)
        Dim podkatalogZdjecia As String = Path.Combine("Zdjecia_zamowien", zamowienieId)
        Dim katalogPelny As String = Path.Combine(PodajKatalogPlikow(projektId), podkatalogZdjecia)
        Dim liczbaZdjec As Integer = PodajLiczbePlikowWKatalogu(katalogPelny)
        liczbaZdjec += 1
        podkatalogZdjecia = Path.Combine(podkatalogZdjecia, String.Format("{0}_{1}.{2}", zamowienieId, liczbaZdjec, extension.Replace(".", "")))
        Return podkatalogZdjecia
    End Function
    <WebMethod()> _
    Public Function ZamowienieZdjecieDodaj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zamowienieId As Integer, ByVal zdjecie As Byte(), ByVal rozszezenie As String, ByVal opisZdjecia As String, ByVal nazwaKamery As String, ByVal czyUkryte As Boolean) As ZamowienieZdjecieDodajWynik
        Dim wynik As New ZamowienieZdjecieDodajWynik

        Using cnn As New ConnectionSuperPaker

            If zdjecie Is Nothing Then
                wynik.status = -1
                wynik.status_opis = "Nie przesłano danych zdjęcia"
            End If

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            'Zapisuje tylko podkatalog bo nie chce pokazywac gdzie na dysku przechowywujemy dane
            Dim podkatalogINazwaZdjecia As String

            Try
                podkatalogINazwaZdjecia = PodajPodkatalogINazweZdjecia(projekt_id, zamowienieId, rozszezenie)
            Catch ex As Exception
                wynik.status = -1
                wynik.status_opis = "Błąd zapisu zdjęcia"
                Return wynik
            End Try

            Dim wynikZapiszPlik As ZapiszPlikWynik
            wynikZapiszPlik = ZapiszPlik(sesja, projekt_id, podkatalogINazwaZdjecia, zdjecie)

            If wynikZapiszPlik.status = -1 Then
                wynik.status = wynikZapiszPlik.status
                wynik.status_opis = wynikZapiszPlik.status_opis
                Return wynik
            End If

            Dim rozmiarBajtow As Integer = zdjecie.Length

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@SCIEZKA", podkatalogINazwaZdjecia)
            htParams.Add("@CZY_UKRYTE", czyUkryte)
            htParams.Add("@WIELKOSC", rozmiarBajtow)
            htParams.Add("@NAZWA_KAMERY", nazwaKamery)
            htParams.Add("@OPIS_ZDJECIA", opisZdjecia)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_ZDJECIE_DODAJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ZamowienieZdjeciaListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zamowienieId As Integer) As ZamowienieZdjeciaListaPobierzWynik
        Dim wynik As New ZamowienieZdjeciaListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_ZDJECIA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieZdjecieEdytuj(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zdjecieId As Integer, ByVal ukryj As Boolean, ByVal usun As Boolean) As ZamowienieZdjecieEdytujWynik
        Dim wynik As New ZamowienieZdjecieEdytujWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ZDJECIE_ID_IN", zdjecieId)
            htParams.Add("@UKRYJ", ukryj)
            htParams.Add("@USUN", usun)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZAMOWIENIE_ZDJECIE_EDYTUJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ZamowienieZdjeciaObrazListaPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal zamowienieId As Integer) As ZamowienieZdjeciaObrazListaPobierzWynik
        Dim wynik As New ZamowienieZdjeciaObrazListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projekt_id)
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_ZDJECIA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet

                Dim KOLUMNA_ZDJECIA As String = "ZDJECIE_BIN"
                Dim KOLUMNA_SCIEZKA As String = "SCIEZKA"

                If Not wynik.dane Is Nothing AndAlso wynik.dane.Tables.Count > 0 And wynik.dane.Tables(0).Columns.Contains(KOLUMNA_ZDJECIA) = False AndAlso wynik.dane.Tables(0).Columns.Contains(KOLUMNA_SCIEZKA) = True Then
                    wynik.dane.Tables(0).Columns.Add(KOLUMNA_ZDJECIA, GetType(Byte()))

                    For Each zdjecieRow As DataRow In wynik.dane.Tables(0).Rows
                        Dim sciezka As String = zdjecieRow(KOLUMNA_SCIEZKA)
                        Dim pobieranieZdjeciaWynik As PobierzPlikWynik = PobierzPlik(sesja, projekt_id, sciezka)

                        If pobieranieZdjeciaWynik.status = 0 Then
                            Dim plik As Byte() = pobieranieZdjeciaWynik.plik
                            zdjecieRow(KOLUMNA_ZDJECIA) = plik
                        End If
                    Next
                End If

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function


#End Region

#Region "Operacje dyskowe"
    Private Function PodajKatalogPlikow(ByVal projektId As String) As String
        Return String.Format(CONST_SCIEZKA_PLIKOW, projektId)
    End Function
    Private Sub StworzKatalog(ByVal sciezka As String)
        Directory.CreateDirectory(Path.GetDirectoryName(sciezka))
    End Sub
    Private Function PobierzPlik(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikSciezka As String) As PobierzPlikWynik
        Dim wynik As New PobierzPlikWynik
        Try
            Dim sciezkaPliku As String = Path.Combine(PodajKatalogPlikow(projektId), plikSciezka)

            If File.Exists(sciezkaPliku) = False Then
                wynik.status = -1
                wynik.status_opis = "Plik nie istnieje"
                Return wynik
            End If

            Dim plik() As Byte
            plik = File.ReadAllBytes(sciezkaPliku)

            wynik.status = 0
            wynik.status_opis = "Pobrano plik"
            wynik.plik = plik
            Return wynik
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = ex.Message
            Return wynik
        End Try
    End Function
    Private Function UsunPlik(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikSciezka As String) As UsunPlikWynik
        Dim wynik As New UsunPlikWynik
        Try
            Dim sciezkaPliku As String = Path.Combine(PodajKatalogPlikow(projektId), plikSciezka)

            If File.Exists(sciezkaPliku) = False Then
                wynik.status = -1
                wynik.status_opis = "Plik nie istnieje"
                Return wynik
            End If

            File.Delete(sciezkaPliku)
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = ex.Message
            Return wynik
        End Try

        wynik.status = 0
        wynik.status_opis = "Usunięto plik"
        Return wynik
    End Function
    Private Function ZapiszPlik(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal plikSciezka As String, ByVal plikDane As Byte()) As ZapiszPlikWynik
        Dim wynik As New ZapiszPlikWynik
        Try
            Dim sciezkaPliku As String = Path.Combine(PodajKatalogPlikow(projektId), plikSciezka)
            StworzKatalog(sciezkaPliku)

            File.WriteAllBytes(sciezkaPliku, plikDane)
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = "Błąd zapisu pliku" 'ex.Message
            Return wynik
        End Try

        wynik.status = 0
        wynik.status_opis = "Zapisano plik"
        Return wynik
    End Function
#End Region

#Region "Zwroty"

    <WebMethod()> _
    Public Function ZamowienieZwrotHistDanePobierz(ByVal sesja As Byte(), ByVal zamowienieId As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZAMOWIENIE_ZWROTY_HIST_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotNiezidentyfikowanyListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal statusyDS As DataSet) As ZwrotListaWynik

        Dim wynik As New ZwrotListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableStatusy As DataTable = New DataTable
        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STATUSY_LISTA_IN", dTable)
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_NIEZIDENTYFIKOWANY_LISTA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotNiezidentyfikowanyStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer,
                                      ByVal sortPo As String,
                                      ByVal sortAsc As Boolean
                                           ) As ZwrotStronaWynik

        Dim wynik As New ZwrotStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                htParams.Add("@STRONA_NR_IN", stronaNumer)
                htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                htParams.Add("@SORT_PO_IN", sortPo)
                htParams.Add("@SORT_KIERUNEK_IN", sortAsc)
                htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_ZWROT_NIEZIDENTYFIKOWANY_STRONA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                If procRes.Status = 0 Then
                    wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                End If
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotNiezidentyfikowanyDanePobierz(ByVal sesja As Byte(), ByVal zwrotId As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZWROT_ID_IN", zwrotId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_NIEZIDENTYFIKOWANY_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal statusyDS As DataSet) As ZwrotListaWynik

        Dim wynik As New ZwrotListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim dTableStatusy As DataTable = New DataTable
        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_ROZPOZNANY_LISTA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyStronaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal stronaNumer As Integer,
                                      ByVal stronaWielkosc As Integer,
                                      ByVal sortPo As String,
                                      ByVal sortAsc As Boolean,
                                      ByVal statusyDS As DataSet) As ZwrotStronaWynik

        Dim wynik As New ZwrotStronaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim totalIloscWierszy As Integer = 0

        Dim dTableStatusy As DataTable = New DataTable
        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy
            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STRONA_NR_IN", stronaNumer)
                    htParams.Add("@STRONA_WIELKOSC_IN", stronaWielkosc)
                    htParams.Add("@SORT_PO_IN", sortPo)
                    htParams.Add("@SORT_KIERUNEK_IN", sortAsc)
                    htParams.Add("@OUT_ILOSC_WIERSZY_TOTAL", totalIloscWierszy)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_ZWROT_ROZPOZNANY_STRONA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    If procRes.Status = 0 Then
                        wynik.totalIloscWierszy = htParams("@OUT_ILOSC_WIERSZY_TOTAL")
                    End If
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If
        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyDanePobierz(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zwrotId As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZWROT_ID_IN", zwrotId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_DANE_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyDodajZmien(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zwrotId As Integer,
                                        ByVal adresId As Integer, ByVal imie As String, ByVal nazwisko As String, ByVal nazwa As String, ByVal firma As String, _
                                        ByVal ulica As String, ByVal nrDomu As String, ByVal nrMieszkania As String, ByVal urzadPocztowy As String, ByVal krajRegion As String, _
                                        ByVal wojewodztwo As String, ByVal kodPocztowy As String, ByVal miasto As String, ByVal kraj As String, _
                                        ByVal e_mail As String, ByVal telefon As String, ByVal telefonKomorkowy As String, ByVal fax As String, _
                                        ByVal nip As String, ByVal numerVat As String, ByVal dataOd As DateTime, ByVal dataDo As DateTime?, _
                                        ByVal przewoznikId As Integer, przyczynaId As Integer, zwrotListNr As String, uwagi As String, _
                                        ByVal dataNadania As DateTime, ByVal rejestracjaData As DateTime?) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZWROT_ID_IN", zwrotId)
            htParams.Add("@ZWROT_NADAWCA_ADRES_ID_IN", NZ(adresId, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_IMIE_IN", NZ(imie, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NAZWISKO_IN", NZ(nazwisko, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NAZWA_IN", NZ(nazwa, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_FIRMA_IN", NZ(firma, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_ULICA_IN", NZ(ulica, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NR_DOMU_IN", NZ(nrDomu, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NR_MIESZKANIA_IN", NZ(nrMieszkania, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_URZAD_POCZTOWY_IN", NZ(urzadPocztowy, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_KRAJ_REGION_IN", NZ(krajRegion, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_WOJEWODZTWO_IN", NZ(wojewodztwo, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_KOD_POCZTOWY_IN", NZ(kodPocztowy, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_MIASTO_IN", NZ(miasto, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_KRAJ_IN", NZ(kraj, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_E_MAIL_IN", NZ(e_mail, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_TELEFON_IN", NZ(telefon, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_TELEFON_KOM_IN", NZ(telefonKomorkowy, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_FAX_IN", NZ(fax, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NIP_IN", NZ(nip, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_NUMER_VAT_IN", NZ(numerVat, DBNull.Value))
            htParams.Add("@ZWROT_NADAWCA_DATA_OD_IN", dataOd)
            htParams.Add("@ZWROT_NADAWCA_DATA_DO_IN", dataDo) 'IIf(dataDo.HasValue, dataDo, DBNull.Value))
            htParams.Add("@ZWROT_PRZEWOZNIK_ID_IN", przewoznikId)
            htParams.Add("@ZWROT_PRZYCZYNA_ID_IN", przyczynaId)
            htParams.Add("@ZWROT_LIST_NR_IN", zwrotListNr)
            htParams.Add("@UWAGI_IN", uwagi)
            htParams.Add("@ZWROT_REJESTRACJA_DATA_IN", rejestracjaData) 'IIf(rejestracjaData.HasValue, rejestracjaData, DBNull.Value))
            htParams.Add("@ZWROT_DATA_NADANIA_IN", dataNadania)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_DODAJ_ZMIEN", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyPozDodajZmien(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zamowieniePozId As Integer, ByVal zwrotId As Integer, ByVal zwrotPozStanId As Integer, ByVal ilosc As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZAMOWIENIE_POZYCJA_ID_IN", zamowieniePozId)
            htParams.Add("@ZWROT_ID_IN", zwrotId)
            htParams.Add("@ZWROT_POZYCJA_STAN_ID_IN", zwrotPozStanId)
            htParams.Add("@ILOSC_IN", ilosc)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_POZYCJA_DODAJ_ZMIEN", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotPrzewoznikListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_PRZEWOZNIK_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotPrzyczynaListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_PRZYCZYNA_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotStanPozycjaListaPobierz(ByVal sesja As Byte(), projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_POZYCJA_STAN_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyZarejestruj(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zwrotId As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieId)
            htParams.Add("@ZWROT_ID_IN", zwrotId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_ZAREJESTRUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyRozliczenieZapisz(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal zwrotId As Integer,
                                        ByVal przelewData As DateTime?, ByVal fakturaKorektaListNr As String, korektaPozycje As DataSet, ByVal zwrotPrzyczynaID As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        If korektaPozycje.Tables.Count > 0 Then
            dTable = korektaPozycje.Tables(0).Copy


            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@ZWROT_ID_IN", zwrotId)
                htParams.Add("@TableFKorektaPozycja_in", dTable)
                htParams.Add("@ZWROT_PRZELEW_DATA_IN", przelewData) 'IIf(fakturaKorektaData.HasValue, fakturaKorektaData, DBNull.Value))
                htParams.Add("@FAKTURA_KOREKTA_LIST_NR_IN", NZ(fakturaKorektaListNr, DBNull.Value))
                htParams.Add("@ZWROT_PRZYCZYNA_ID_IN", zwrotPrzyczynaID)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_ROZLICZENIE_ZAPISZ", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane pozycji korekty"
        End If


        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyRozlicz(ByVal sesja As Byte(), ByVal zwrotId As Integer) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZWROT_ID_IN", zwrotId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_ROZLICZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotDoFakturyKorektyListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal zamowienieID As Integer
                                     ) As ZwrotListaWynik

        Dim wynik As New ZwrotListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                ' htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZAMOWIENIE_ID_IN", zamowienieID)

                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZAMOWIENIE_FKOREKTA_V_POZ_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotKorektaFakturyDokumentDanePobierz(ByVal Projekt As Integer, ByVal sesja As Byte(),
                                      ByVal zamowienieId As Integer, ByVal ZwrotId As Integer, ByVal DokumentData As DateTime,
                                      ByVal dokumentTypId As Integer, ByVal dokumentTypGrupaId As Integer) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dokumentId As Integer = 0
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()

            Try
                cnn.Open()
                htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
                htParams.Add("@ZWROT_ID", ZwrotId)
                htParams.Add("@DOKUMENT_DATA", DokumentData)
                htParams.Add("@DOKUMENT_TYP_ID", dokumentTypId)
                htParams.Add("@DOKUMENT_TYP_GRUPA_ID", dokumentTypGrupaId)
                htParams.Add("@OUT_DOKUMENT_ID", dokumentId)
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_FKOREKTA_GENERUJ", htParams, dSet)
                'dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotZamowienieListaPobierz(ByVal sesja As Byte(),
                                      ByVal projektId As Integer,
                                      ByVal zamowienieStatusId As Integer,
                                      ByVal wyszukajPoId As Integer,
                                      ByVal wyszukajText As String,
                                      ByVal statusyDS As DataSet) As ZamowienieListaWynik

        Dim wynik As New ZamowienieListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim dTableStatusy As DataTable = New DataTable

        If statusyDS.Tables.Count > 0 Then
            dTableStatusy = statusyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                Try
                    cnn.Open()
                    htParams.Add("@PROJEKT_ID_IN", projektId)
                    'htParams.Add("@ZAMOWIENIE_STATUS_ID_IN", zamowienieStatusId)
                    htParams.Add("@WYSZUKAJ_TYP_ID_IN", wyszukajPoId)
                    htParams.Add("@WYSZUKAJ_TEKST_IN", wyszukajText)
                    htParams.Add("@STATUSY_LISTA_IN", dTableStatusy)
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_ZAMOWIENIE_LISTA_POBIERZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane filtra statusów"
        End If

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotStatusyListaPobierz(ByVal sesja As Byte(), projektId As Integer, filtr As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@FILTR_STATUSOW_IN", filtr)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_STATUSY_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotPrzewoznikZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal przewoznik_id As Integer, ByVal przewoznikNazwa As String, ByVal przewoznikOpis As String, ByVal imie As String, _
                                           ByVal nazwisko As String, ByVal nazwa As String, ByVal firma As String, ByVal ulica As String, _
                                           ByVal nr_domu As String, ByVal nr_mieszkania As String, ByVal urzad_pocztowy As String, ByVal kraj_region As String, _
                                           ByVal wojewodztwo As String, ByVal kod_pocztowy As String, ByVal miasto As String, ByVal kraj As String, _
                                           ByVal e_mail As String, ByVal telefon As String, ByVal telefon_komorkowy As String, ByVal fax As String, _
                                           ByVal nip As String, ByVal numer_vat As String) As ZwrotPrzewoznikZapiszWynik

        Dim wynik As New ZwrotPrzewoznikZapiszWynik

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@PRZEWOZNIK_NAZWA", NZ(przewoznikNazwa, DBNull.Value))
            htParams.Add("@PRZEWOZNIK_OPIS", NZ(przewoznikOpis, DBNull.Value))
            htParams.Add("@IMIE", NZ(imie, DBNull.Value))
            htParams.Add("@NAZWISKO", NZ(nazwisko, DBNull.Value))
            htParams.Add("@NAZWA", NZ(nazwa, DBNull.Value))
            htParams.Add("@FIRMA", NZ(firma, DBNull.Value))
            htParams.Add("@ULICA", NZ(ulica, DBNull.Value))
            htParams.Add("@NR_DOMU", NZ(nr_domu, DBNull.Value))
            htParams.Add("@NR_MIESZKANIA", NZ(nr_mieszkania, DBNull.Value))
            htParams.Add("@URZAD_POCZTOWY", NZ(urzad_pocztowy, DBNull.Value))
            htParams.Add("@KRAJ_REGION", NZ(kraj_region, DBNull.Value))
            htParams.Add("@WOJEWODZTWO", NZ(wojewodztwo, DBNull.Value))
            htParams.Add("@KOD_POCZTOWY", NZ(kod_pocztowy, DBNull.Value))
            htParams.Add("@MIASTO", NZ(miasto, DBNull.Value))
            htParams.Add("@KRAJ", NZ(kraj, DBNull.Value))
            htParams.Add("@E_MAIL", NZ(e_mail, DBNull.Value))
            htParams.Add("@TELEFON", NZ(telefon, DBNull.Value))
            htParams.Add("@TELEFON_KOMORKOWY", NZ(telefon_komorkowy, DBNull.Value))
            htParams.Add("@FAX", NZ(fax, DBNull.Value))
            htParams.Add("@NIP", NZ(nip, DBNull.Value))
            htParams.Add("@NUMER_VAT", NZ(numer_vat, DBNull.Value))
            htParams.Add("@OUT_ZWROT_PRZEWOZNIK_ID", przewoznik_id)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_ZWROT_PRZEWOZNIK_DODAJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally

                wynik.przewoznik_id = NZ(htParams("@OUT_ZWROT_PRZEWOZNIK_ID"), -1)
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using


        Return wynik
    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyDaneSzczegolyAPIPobierz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal zwrotId As Integer) As ZwrotRozpoznanyDaneSzczegolyAPIWynik

        Dim wynik As New ZwrotRozpoznanyDaneSzczegolyAPIWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektID)
            htParams.Add("@ZWROT_ID_IN", zwrotId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_ROZPOZNANY_DANE_POBIERZ_READONLY", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function ZwrotRozpoznanyListaAPIPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                      ByVal dataRejestracjiZwrotu As DateTime) As ZwrotRozpoznanyListaAPIWynik

        Dim wynik As New ZwrotRozpoznanyListaAPIWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet


        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            Try
                cnn.Open()
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@DATA_REJESTRACJI", IIf(dataRejestracjiZwrotu = DateTime.MinValue, DBNull.Value, dataRejestracjiZwrotu))

                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_ZWROT_ROZPOZNANY_LISTA_POBIERZ_PO_DACIE", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function
    <WebMethod()> _
    Public Function FakturaKorektaGeneruj(ByVal sesja As Byte(), ByVal zamowienieId As Integer, ByVal dataKorekty As DateTime, ByVal nabywcaNazwa As String, ByVal nabywcaOsoba As String, ByVal nabywcaKodPocztowy As String, ByVal nabywcaMiasto As String, ByVal nabywcaUlica As String, ByVal nabywcaNip As String) As FakturaKorektaGenerujWynik

        Dim wynik As New FakturaKorektaGenerujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim fvSkorygowanaOut As Integer = -1

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@ZAMOWIENIE_ID", zamowienieId)
            htParams.Add("@DOK_FV_ZMIANA_DATA", dataKorekty)
            htParams.Add("@NABYWCA_NAZWA", nabywcaNazwa)
            htParams.Add("@NABYWCA_OSOBA", nabywcaOsoba)
            htParams.Add("@NABYWCA_KOD_POCZTOWY", nabywcaKodPocztowy)
            htParams.Add("@NABYWCA_MIASTO", nabywcaMiasto)
            htParams.Add("@NABYWCA_ULICA", nabywcaUlica)
            htParams.Add("@NABYWCA_NIP", nabywcaNip)
            htParams.Add("@OUT_DOK_FV_ZMIANA_ID", fvSkorygowanaOut)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_DOK_ALL_FZMIANA_GENERUJ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                fvSkorygowanaOut = NZ(htParams("@OUT_DOK_FV_ZMIANA_ID"), -1)
                If wynik.status = 0 AndAlso fvSkorygowanaOut > 0 Then
                    wynik.fakturaSkorygowanaId = fvSkorygowanaOut
                End If

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik

    End Function

    <WebMethod()> _
    Public Function ZwrotRozpoznanyZPozycjamiDodajZmien(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieNumer As String, ByVal zwrotId As Integer,
                                        ByVal adresId As Integer, ByVal imie As String, ByVal nazwisko As String, ByVal nazwa As String, ByVal firma As String, _
                                        ByVal ulica As String, ByVal nrDomu As String, ByVal nrMieszkania As String, ByVal urzadPocztowy As String, ByVal krajRegion As String, _
                                        ByVal wojewodztwo As String, ByVal kodPocztowy As String, ByVal miasto As String, ByVal kraj As String, _
                                        ByVal e_mail As String, ByVal telefon As String, ByVal telefonKomorkowy As String, ByVal fax As String, _
                                        ByVal nip As String, ByVal numerVat As String, ByVal dataOd As DateTime, ByVal dataDo As DateTime?, _
                                        ByVal przewoznikId As Integer, przyczynaId As Integer, zwrotListNr As String, uwagi As String, _
                                        ByVal dataNadania As DateTime, ByVal rejestracjaData As DateTime?, ByVal pozycjeZwrotuDs As DataSet) As ZwrotDaneWynik

        Dim wynik As New ZwrotDaneWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Dim pozycjeZwrotu As DataTable

        If IsNothing(pozycjeZwrotuDs) = False AndAlso pozycjeZwrotuDs.Tables.Count > 0 Then
            pozycjeZwrotu = pozycjeZwrotuDs.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@KLIENT_NUMER_ZAMOWIENIA_IN", zamowienieNumer)
                htParams.Add("@PROJEKT_ID_IN", projektId)
                htParams.Add("@ZWROT_ID_IN", zwrotId)
                htParams.Add("@ZWROT_NADAWCA_ADRES_ID_IN", NZ(adresId, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_IMIE_IN", NZ(imie, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NAZWISKO_IN", NZ(nazwisko, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NAZWA_IN", NZ(nazwa, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_FIRMA_IN", NZ(firma, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_ULICA_IN", NZ(ulica, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NR_DOMU_IN", NZ(nrDomu, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NR_MIESZKANIA_IN", NZ(nrMieszkania, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_URZAD_POCZTOWY_IN", NZ(urzadPocztowy, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_KRAJ_REGION_IN", NZ(krajRegion, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_WOJEWODZTWO_IN", NZ(wojewodztwo, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_KOD_POCZTOWY_IN", NZ(kodPocztowy, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_MIASTO_IN", NZ(miasto, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_KRAJ_IN", NZ(kraj, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_E_MAIL_IN", NZ(e_mail, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_TELEFON_IN", NZ(telefon, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_TELEFON_KOM_IN", NZ(telefonKomorkowy, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_FAX_IN", NZ(fax, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NIP_IN", NZ(nip, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_NUMER_VAT_IN", NZ(numerVat, DBNull.Value))
                htParams.Add("@ZWROT_NADAWCA_DATA_OD_IN", dataOd)
                htParams.Add("@ZWROT_NADAWCA_DATA_DO_IN", dataDo) 'IIf(dataDo.HasValue, dataDo, DBNull.Value))
                htParams.Add("@ZWROT_PRZEWOZNIK_ID_IN", przewoznikId)
                htParams.Add("@ZWROT_PRZYCZYNA_ID_IN", przyczynaId)
                htParams.Add("@ZWROT_LIST_NR_IN", zwrotListNr)
                htParams.Add("@UWAGI_IN", uwagi)
                htParams.Add("@ZWROT_REJESTRACJA_DATA_IN", rejestracjaData) 'IIf(rejestracjaData.HasValue, rejestracjaData, DBNull.Value))
                htParams.Add("@ZWROT_DATA_NADANIA_IN", dataNadania)
                htParams.Add("@ZWROT_POZYCJE_LISTA", NZ(pozycjeZwrotu, DBNull.Value))



                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_ZWROT_DO_ZAMOWIENIA_ZAPISZ", htParams, dSet)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message
                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using

        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane zwrotu do zapisania"
        End If

        Return wynik

    End Function

#End Region

#Region "Szablon importu"

    <WebMethod()> _
    Public Function SzablonImportuPolaTypyListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer
                                           ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_POLE_TYP_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuSzablonyListaPobierz(ByVal sesja As Byte(), projektId As Integer, ByVal grupaID As Integer,
                                           ByVal SzablonTypID As Integer, ByVal SzablonNazwa As String, ByVal SzablonDomenaNazwa As String) As SzablonImportuSzablonListaWynik

        Dim wynik As New SzablonImportuSzablonListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@GRUPA_ID_IN", grupaID)
            htParams.Add("@SZABLON_TYP_ID_IN", SzablonTypID)
            htParams.Add("@SZABLON_ID_IN", SzablonNazwa)
            htParams.Add("@DOMENA_NAZWA_IN", SzablonDomenaNazwa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_SZABLON_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuSzablonPozycjaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                           ByVal SzablonID As Integer) As SzablonImportuSzablonPozycjaListaWynik

        Dim wynik As New SzablonImportuSzablonPozycjaListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@SZABLON_ID_IN", SzablonID)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_SZABLON_POZYCJA_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuSzablonTypListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer
                                                             ) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_SZABLON_TYP_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuDomenaTypListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                                        ByVal domenaNazwa As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DOMENA_TYP_NAZWA_IN", domenaNazwa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_DOMENA_TYP_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuDomenaListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                                        ByVal DomenaTypID As Integer, ByVal DomenaNazwa As String) As SlownikListaWynik

        Dim wynik As New SlownikListaWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DOMENA_TYP_ID_IN", DomenaTypID)
            htParams.Add("@DOMENA_NAZWA_IN", DomenaNazwa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_DOMENA_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuDomenaElementListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer,
                                                         ByVal DomenaID As Integer) As SlownikItemsWynik

        Dim wynik As New SlownikItemsWynik
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DOMENA_ID_IN", DomenaID)
            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDS(cnn, sesja, "UP_SI_DOMENA_ELEMENT_LISTA_POBIERZ", htParams, dSet)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuDodajEdytuj(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal grupaIdIn As Integer, ByVal szablonId As Integer,
                                          ByVal szabonNazwa As String, ByVal szablonTypId As Integer, ByVal szablonDomenaId As Integer,
                                          ByVal szablonLnNaglowekStart As Integer, ByVal szablonLnDaneStart As Integer, ByVal szablonLnDaneWielkosc As Integer,
                                          ByVal szablonOpis As String, ByVal szablonPozycjeDS As DataSet) As SzablonImportuSzablonDodajEdytujWynik
        Dim wynik As New SzablonImportuSzablonDodajEdytujWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim pozycjeImportuDT As DataTable = New DataTable

        If szablonPozycjeDS.Tables.Count > 0 Then
            pozycjeImportuDT = szablonPozycjeDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID_IN", projektID)
                htParams.Add("@GRUPA_ID_IN", grupaIdIn)
                htParams.Add("@SZABLON_ID_IN", szablonId)
                htParams.Add("@SZABLON_TYP_ID_IN", szablonTypId)
                htParams.Add("@SZABLON_NAZWA_IN", szabonNazwa)
                htParams.Add("@SZABLON_DOMENA_ID_IN", szablonDomenaId)
                htParams.Add("@SZABLON_LN_NAGLOWEK_START_IN", szablonLnNaglowekStart)
                htParams.Add("@SZABLON_LN_DANE_START_IN", szablonLnDaneStart)
                htParams.Add("@SZABLON_LN_DANE_WIELKOSC_IN", szablonLnDaneWielkosc)
                htParams.Add("@SZABLON_OPIS_IN", szablonOpis)
                htParams.Add("@TableSzablonPozycje_in", pozycjeImportuDT)


                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_SI_SZABLON_DODAJ_ZMIEN", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane imoprtu"
        End If

        Return wynik
    End Function

    <WebMethod()> _
    Public Function SzablonImportuUniewaznij(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal szablonID As Integer, ByVal czyUniewaznic As Integer) As SzablonImportuUniewaznijWynik
        Dim wynik As New SzablonImportuUniewaznijWynik
        Dim czyUniewaznionyOUT As Boolean

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@SZABLON_ID_IN", szablonID)
            htParams.Add("@CZY_UNIEWAZNIC_IN", czyUniewaznic)
            htParams.Add("@OUT_CZY_UNIEWAZNIONY", czyUniewaznionyOUT)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_SI_SZABLON_UNIEWAZNIJ", htParams)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.czyUniewazniony = NZ(htParams("@OUT_CZY_UNIEWAZNIONY"), False)

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function PlikSzablonuPobierz(ByVal sesja As Byte(), ByVal projekt_id As Integer, ByVal szablonNazwa As String) As PobierzPlikWynik
        Dim wynik As New PobierzPlikWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        
        Try
            If szablonNazwa.Length = 0 Then
                wynik.status = -1
                wynik.status_opis = "Nie podano nazwy pliku"
            Else
                Dim sciezka As String = Path.Combine(CONST_SCIEZKA_PLIKOW_SZABLONOW, szablonNazwa)
                wynik = PobierzPlik(sesja, projekt_id, sciezka)
            End If
        Catch ex As Exception
            wynik.status = -1
            wynik.status_opis = String.Format("Pobieranie pliku szablonu :{0}{1}", ex.Message, kontaktIt)
        End Try

        Return wynik
    End Function
#End Region

#Region "Obsługa gratisów"
    <WebMethod()> _
    Public Function ProduktGratisListaZapisz(ByVal sesja As Byte(), ByVal projektID As Integer, ByVal gratisyDS As DataSet) As ProduktGratisListaZapiszWynik
        Dim wynik As New ProduktGratisListaZapiszWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim gratisyPozycjeDT As DataTable
        If gratisyDS.Tables.Count > 0 Then
            gratisyPozycjeDT = gratisyDS.Tables(0).Copy

            Using cnn As New ConnectionSuperPaker

                Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
                procRes.Status = DataAccess.Status.Error
                procRes.Message = ""

                Dim htParams As Hashtable = New Hashtable()
                htParams.Add("@PROJEKT_ID", projektID)
                htParams.Add("@PRODUKT_GRATIS_LISTA", gratisyPozycjeDT)

                Try
                    cnn.Open()
                    procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKT_GRATIS_LISTA_ZAPISZ", htParams, dTable)
                    dSet.Tables.Add(dTable)
                    wynik.dane = dSet
                Catch ex As Exception
                    procRes.Status = -1
                    procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

                Finally
                    wynik.status = procRes.Status
                    wynik.status_opis = procRes.Message

                    If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                        cnn.Close()
                    End If
                End Try

            End Using
        Else
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "Nieprawidłowe dane gratisów do zapisania"
        End If

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ProduktGratisListaPobierz(ByVal sesja As Byte(), ByVal projektId As Integer) As ProduktGratisListaPobierzWynik
        Dim wynik As New ProduktGratisListaPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "UP_PRODUKT_GRATIS_LISTA_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                wynik.dane = dSet
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function
    <WebMethod()> _
    Public Function ZamowienieGratisPobierz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal zamowienieID As Integer) As ZamowienieGratisPobierzWynik
        Dim wynik As New ZamowienieGratisPobierzWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet
        Dim PopupMessage As String = ""
        Dim InformacjeDodatkowe As String = ""
        Dim InformacjeHtml As String = ""
        Dim GrafikaPrawa As Byte() = New Byte() {}
        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID", projektId)
            htParams.Add("@ZAMOWIENIE_ID", zamowienieID)
            htParams.Add("@OUT_POPUP_MESSAGE:-1", PopupMessage)
            htParams.Add("@OUT_INFORMACJE_DODATKOWE:-1", InformacjeDodatkowe)
            htParams.Add("@OUT_HTML:-1", InformacjeHtml)
            htParams.Add("@OUT_GRAFIKA_PRAWA:-1", GrafikaPrawa)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDTOutput(cnn, sesja, "UP_ZAMOWIENIE_INF_DOD_POBIERZ", htParams, dTable)
                dSet.Tables.Add(dTable)
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message
                If procRes.Status = 0 Then
                    wynik.PopupMessage = NZ(htParams("@OUT_POPUP_MESSAGE:-1"), "")
                    wynik.InformacjeDodatkowe = NZ(htParams("@OUT_INFORMACJE_DODATKOWE:-1"), "")
                    wynik.InformacjeHtml = NZ(htParams("@OUT_HTML:-1"), "")
                    wynik.GrafikaPrawa = NZ(htParams("@OUT_GRAFIKA_PRAWA"), "")
                End If
                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

#End Region

#Region "Testery sieci"

    <WebMethod()> _
    Public Function PerformBandwidthTest(ByVal fileSizeMb As Integer) As PerformBandwidthTestWynik
        Dim wynik As New PerformBandwidthTestWynik


        Try
            If fileSizeMb > 10 OrElse fileSizeMb < 1 Then
                wynik.dane = Nothing
                wynik.status = 0
                wynik.status_opis = "PerformBandwidthTest choose fileSizeMb from 1 to 10"
                Return wynik
            End If

            Dim bytes As Integer = fileSizeMb * 1024 * 1024
            Dim data As Char() = New Char(bytes - 1) {}

            wynik.dane = Encoding.ASCII.GetBytes(data)
            wynik.status = 0
            wynik.status_opis = "OK"

        Catch ex As Exception
            wynik.dane = Nothing
            wynik.status = -1
            wynik.status_opis = "PerformBandwidthTest exception: " & ex.Message
        End Try


        Return wynik
    End Function



#End Region

#Region "Helpers"
    Public Function getListaTypeTbl() As DataTable
        Dim listaTypeTable As DataTable = New DataTable
        listaTypeTable.TableName = "tblListaTypeTbl"
        listaTypeTable.Columns.Add([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.ID), Type.GetType("System.Int32"))
        listaTypeTable.Columns.Add([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_INT), Type.GetType("System.Int32"))
        listaTypeTable.Columns.Add([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_NVARCHAR), Type.GetType("System.String"))
        listaTypeTable.Columns.Add([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_BIN), Type.GetType("System.Byte[]"))

        listaTypeTable.Columns([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.ID)).SetOrdinal(0)
        listaTypeTable.Columns([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_INT)).SetOrdinal(1)
        listaTypeTable.Columns([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_NVARCHAR)).SetOrdinal(2)
        listaTypeTable.Columns([Enum].GetName(GetType(ListaTypeEnum), ListaTypeEnum.WARTOSC_BIN)).SetOrdinal(3)
        Return listaTypeTable
    End Function

#End Region


#Region "Dev TEST"

    <WebMethod()> _
    Public Function TestSzukajki(ByVal sesja As Byte(), ByVal filtr As String) As SlownikListaWynik
        Dim wynik As New SlownikListaWynik
        Dim dTable As DataTable = New DataTable
        Dim dSet As DataSet = New DataSet

        Using cnn As New ConnectionSuperPaker

            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@FILTR", filtr)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProcDT(cnn, sesja, "_PROC_CustomerAddress", htParams, dTable)
                dSet.Tables.Add(dTable)
                wynik.dane = dSet
            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)

            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function DokumentSzablonWydrukuGrafikaDodaj(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal dokumentSzablonID As Integer, ByVal dokumentPoleNazwa As String, ByVal obraz As Byte()) As SzablonWydrukuGrafikaDodajWynik
        Dim wynik As New SzablonWydrukuGrafikaDodajWynik
        Using cnn As New ConnectionSuperPaker


            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@DOK_SZABLON_ID_IN", dokumentSzablonID)
            htParams.Add("@DOK_POLE_IN", dokumentPoleNazwa)
            htParams.Add("@IMAGE_DATA_IN", obraz)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_DOK_SZABLON_GRAFIKA_DODAJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

    <WebMethod()> _
    Public Function ProjektLogoZapisz(ByVal sesja As Byte(), ByVal projektId As Integer, ByVal logoProjektu As Byte()) As ProjektLogoDodajWynik
        Dim wynik As New ProjektLogoDodajWynik
        Using cnn As New ConnectionSuperPaker


            Dim procRes As DataAccess.ProcedureResult = New DataAccess.ProcedureResult
            procRes.Status = DataAccess.Status.Error
            procRes.Message = ""

            Dim htParams As Hashtable = New Hashtable()
            htParams.Add("@PROJEKT_ID_IN", projektId)
            htParams.Add("@IMAGE_DATA_IN", logoProjektu)

            Try
                cnn.Open()
                procRes = DataAccess.Helpers.ExecuteProc(cnn, sesja, "UP_PROJEKT_IMG_DODAJ", htParams)

            Catch ex As Exception
                procRes.Status = -1
                procRes.Message = String.Format("Database Communication Error:{0}{1}", ex.Message, kontaktIt)
            Finally
                wynik.status = procRes.Status
                wynik.status_opis = procRes.Message

                If IsNothing(cnn.SqlTran) And Not IsNothing(cnn.SqlConn) Then
                    cnn.Close()
                End If
            End Try

        End Using

        Return wynik
    End Function

#End Region


End Class