#Region "Imports"
Imports System.Xml
#End Region

Public Class Komunikat
    Public id As Integer 'identyfikator komunikatu (aplikacja potwierdza przy jego użyciu odbiór komunikatu)
    Public status As Integer '0 - ok, 1 - uwaga, -1 - błąd
    Public status_opis As String 'opis słowny wyniku wywołania (w przypdaku powodzenia często niepokazywany)
End Class

'Logowanie
Public Class ZalogujWynik
    Public sesja As Byte() 'zwracany identyfikator sesjii
    Public uzytkownik_id As Integer
    Public uzytkownik As String
    Public telefon As String
    Public status As Integer
    Public status_opis As String
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
    Public czy_pierwszy As Integer
    Public adresPolaczenia As String
    Public kodJezyka As String
End Class

Public Class ZmienHasloWynik
    Public status As Integer
    Public status_opis As String
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

'Grupy
Public Class UzytkownikGrupaProjektowaListaPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class UzytkownikGrupaProjektowaDanePobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class UzytkownikGrupaProjektowaZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public grupaId As Integer
    Public dane As DataSet
End Class


Public Class UzytkownikGrupaProjektowa
    Public grupaId As Integer
    Public nazwa As String
    Public opis As String
    Public UzytkownikListaId() As Integer
    Public ProjektListaId() As Integer

End Class

'Uzytkownik
Public Class UzytkownikListaWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class UzytkownikInfoWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class UzytkownikNotyfikacjeEmailDanePobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class UzytkownikZapiszWynik
    Public uzytkownikId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class UzytkownikEdytujWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class UzytkownikNotyfikacjeEmailDaneZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class

'Magazyny
Public Class MagazynyListaWynik
    Public dane As DataSet
    Public magazyny_licznik As Integer
    Public status As Integer
    Public status_opis As String
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

'Magazyny wirtualne
Public Class MagazynWirtualnyListaWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class MagazynWirtualnyDodajWynik
    Public magazyn_wirtualny_id As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class MagazynWirtualnyEdytujWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class MagazynWirtualnyMMDodajWynik
    Public status As Integer
    Public status_opis As String
End Class

'Zamowienia
Public Class ZamowienieZlozWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

'Slowiniki
Public Class SlownikListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class ZamowienieKompletacjaMasowaListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class SlownikItemsWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class AwizoEdytujWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public project_prefix As String
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class AwizoZapiszWynik
    Public awizoId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class AwizoIskaWyslijPonownieWynik
    Public awizoId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class OsobyKontaktoweDodajWynik
    Public osobaId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class AwizoAnulujWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ProduktPrefixPobierzWynik
    Public prefix As String
    Public status As Integer
    Public status_opis As String
End Class

Public Class ProduktyMasowaEdycjaPolaListaPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ProduktyMasowaEdycjaZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

'Dokumenty
Public Class Dokumenty
    Implements ICollection

    Public CollectionName As String
    Private dokArray As New ArrayList()

    Default Public ReadOnly Property Item(index As Integer) As Dokument
        Get
            Return DirectCast(dokArray(index), Dokument)
        End Get
    End Property

    Public Sub CopyTo(a As Array, index As Integer) Implements System.Collections.ICollection.CopyTo
        dokArray.CopyTo(a, index)
    End Sub

    Public ReadOnly Property Count() As Integer Implements System.Collections.ICollection.Count
        Get
            Return dokArray.Count
        End Get
    End Property

    Public ReadOnly Property SyncRoot() As Object Implements System.Collections.ICollection.SyncRoot
        Get
            Return Me
        End Get
    End Property

    Public ReadOnly Property IsSynchronized() As Boolean Implements System.Collections.ICollection.IsSynchronized
        Get
            Return False
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator Implements System.Collections.ICollection.GetEnumerator
        Return dokArray.GetEnumerator()
    End Function

    Public Sub Add(newDokument As Dokument)
        dokArray.Add(newDokument)
    End Sub

End Class

Public Class Dokument
    Public DokumentDane As Byte()
    Public ID As String
    Public Sub New()
    End Sub
    Public Sub New(id As String, dane As Byte())
        DokumentDane = dane
        id = id
    End Sub
End Class

'Blokada
Public Class BlokadaUsunWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class BlokadaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

'StausWynik
Public Class StatusWynik
    Public status As Integer
    Public status_opis As String
End Class

'Zamowienia
Public Class ZamowienieListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZamowienieKompletacjaWydrukiHurtOznaczWydrukowaneListaZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class


Public Class ZamowienieStronaWynik
    Public status As Integer
    Public status_opis As String
    Public totalIloscWierszy As Integer
    Public dane As DataSet
End Class

Public Class ZamowienieDaneWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class


Public Class ZamowienieDaneDodatkoweWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZamowienieParagonBlednyDanePobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZamowienieEdytujWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieZapiszWynik
    Public zamowienieId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieStatusListaWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieEdycjaStatusWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieStatusListaXmlWynik
    Public dane As DataSet
    Public xml As String
    Public status As Integer
    Public status_opis As String
End Class

'Kompletacja
Public Class KompletacjaPaczkiWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class KompletacjaPunktOdbioruWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class KompletacjaUslugiDodatkowePobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class KompletacjaCrossSellDaneKurierPobierzPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class DostawcaEdytujWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public klient_id As Integer
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class DostawcyZapiszWynik
    Public dostawcaId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class ProduktDodajWynik
    Public produktId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class ProduktDaneLogistyczneDodajWynik
    Public produktId As Integer
    Public status As Integer
    Public status_opis As String
End Class

Public Class DokumentWZNumerWynik
    Public dokumentWZNumer As String
    Public status As Integer
    Public status_opis As String
End Class

Public Class DokumentPlikDodajWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class SzablonWydrukuPlikDodajWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class SzablonWydrukuPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class DokumentPlikPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class DokumentInfoPobierzWynik
    Public dokument_dane As String
    Public status As Integer
    Public status_opis As String
End Class

'Raporty
Public Class RaportWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class RaportXMLWynik
    Public status As Integer
    Public status_opis As String
    Public xml As String
    Public dane As DataSet
End Class

'Rozliczenia
Public Class RozliczeniaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class RozliczeniaZamknijWynik
    Public status As Integer
    Public status_opis As String
End Class

'Statystyki
Public Class StatystykiProjektWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class StatystykiWydajnoscWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class StatystykiWydajnoscStronaWynik
    Public status As Integer
    Public status_opis As String
    Public totalIloscWierszy As Integer
    Public dane As DataSet
End Class

Public Class DrukarkaStawkaVATListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class DrukarkaDaneFiskalneZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class SprawdzVatParagonuWynik
    Public roznica As Double
    Public status As Integer
    Public status_opis As String
End Class

'Ustawienia
Public Class ProjektUstawieniaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ProjektUstawieniaSprawdzDostepnoscKonfiguracji
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class StanowiskoUstawieniaPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class
Public Class UstawieniaProjektPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class
Public Class StanowiskoUstawienieDodajEdytujWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class StanowiskoUstawienieUsunWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class PobierzPlikWynik
    Public plik() As Byte
    Public status As Integer
    Public status_opis As String
End Class
Public Class ZapiszPlikWynik
    Public status As Integer
    Public status_opis As String
End Class
Public Class UsunPlikWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowinenieAtrybutyEdycjaZapiszWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowinenieFvDoParagonuGenerujWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieZdjecieDodajWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieZdjeciaListaPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class
Public Class ZamowienieZdjeciaObrazListaPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieZdjecieEdytujWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowinenieUwagiEdycjaZapiszWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

'Zwroty
Public Class ZwrotListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZwrotRozpoznanyListaAPIWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZwrotStronaWynik
    Public status As Integer
    Public status_opis As String
    Public totalIloscWierszy As Integer
    Public dane As DataSet
End Class

Public Class ZwrotDaneWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class FakturaKorektaGenerujWynik
    Public fakturaSkorygowanaId As Integer
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class



Public Class ZwrotRozpoznanyDaneSzczegolyAPIWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZwrotEdytujWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZwrotPrzewoznikZapiszWynik
    Public przewoznik_id As Integer
    Public status As Integer
    Public status_opis As String
End Class


Public Class AdresEdycjaZapiszWynik

    Public status As Integer
    Public status_opis As String
End Class

Public Class AdresEdycjaWarunekSprawdzWynik
    Public OutMoznaEdytowac As Integer
    Public OutMoznaEdytowacOpis As String
    Public status As Integer
    Public status_opis As String
End Class

Public Class ProduktEdytujWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class ProduktEdycjaZapiszWynik
    Public produktId As Integer
    Public status As Integer
    Public status_opis As String
End Class
Public Class ProduktyListaXMLWynik
    Public status As Integer
    Public status_opis As String
    Public xml As String
    Public dane As DataSet
End Class
Public Class ProduktyStronaWynik
    Public status As Integer
    Public status_opis As String
    Public totalIloscWierszy As Integer
    Public dane As DataSet
End Class
Public Class ZamowieniaAbcdataDanePobierzWynik
    Public agregat_id As Integer
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class StanMagazynowyPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class StanMagazynowyUszkodzonePobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class StanMagazynowyProduktListaPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
Public Class StanMagazynowyMagazynWirtualnyPobierzWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ZamowienieListPrzewozowyZmienZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ZamowienieParagonBlednyDanePoprawZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class
Public Class ZamowieniePunktOdbioruZmienZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class
Public Class ZamowienieFakturaEnovaZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public id As Integer
End Class
Public Class ZamowienieDokumentZewnetrznyDaneZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public id As Integer
End Class
#Region "JPK"
Public Class UpoResponse
    Public Code As Integer
    Public Description As String
    Public Details As String
    Public Upo As String
    Public Timestamp As String
End Class

Public Class SaveJpkExecution
    Public upoResponse As UpoResponse
    Public fileXml As String
    Public refNum As String
End Class
#End Region

#Region "Import reczny zamowien"
Public Class ImpKontrahentListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ImpKontrahentZapiszWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ImpZamowienieListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ImpZamowieniePozycjaListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ImpKontrahentIdKlientaPobierzWynik
    Public klientID As String
    Public status As Integer
    Public status_opis As String
End Class

Public Class AdresDlaKoduListaWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ImpZamowienieZapiszWynik
    Public zamowienie As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class ImpZamowienieUsunWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ImpPakietUsunWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class ImpZamowieniePrzeniesDoSTGWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class ImpZamowienieUstawStatusWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class
#End Region
#Region "Szablony importu"
Public Class SzablonImportuSzablonListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class SzablonImportuSzablonPozycjaListaWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class SzablonImportuSzablonDodajEdytujWynik
    Public status As Integer
    Public status_opis As String
    Public dane As DataSet
End Class

Public Class SzablonImportuUniewaznijWynik
    Public czyUniewazniony As Boolean
    Public status_opis As String
    Public status As Integer
End Class
Public Class ProduktGratisListaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class
Public Class ProduktGratisListaZapiszWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class PerformBandwidthTestWynik
    Public dane As Byte()
    Public status As Integer
    Public status_opis As String
End Class


Public Class ZamowienieGratisPobierzWynik
    Public dane As DataSet
    Public PopupMessage As String
    Public InformacjeDodatkowe As String
    Public InformacjeHtml As String
    Public GrafikaPrawa As Byte()
    Public status As Integer
    Public status_opis As String
End Class

Public Class AwizoDostawaZrodloListaWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
    Public komunikaty() As Komunikat 'zaległe komunikaty dla użytkownika
End Class

Public Class MagazynyWirtualneListaPobierzWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class AwizoStronaWynik
    Public status As Integer
    Public status_opis As String
    Public totalIloscWierszy As Integer
    Public dane As DataSet
End Class
#End Region

#Region "Rozliczenia"
Public Class RozliczeniePlikCODRodzajListaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePlikCODDodajWynik
    Public dane As DataSet
    Public status As Integer
    Public status_opis As String
End Class

Public Class RozliczeniePlikCODStatusListaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePlikCODListaPlikowPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class


Public Class RozliczeniePrzelewListaPlikowOczekujacychPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class


Public Class RozliczeniePlikCODPozycjaStanListaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class
Public Class RozliczeniePlikCODWalidacjaSlownikListaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePlikCODPlikSzczegolyPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePrzelewPlikOczekujacySzczegolyPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class
Public Class RozliczeniePlikCODPlikWalidacjaPobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class
Public Class RozliczeniePlikCODOdrzucPlikWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePlikCODDecyzjeZapiszWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class
Public Class RozliczeniePrzelewPlikDodajZmienWynikWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class


Public Class RozliczeniePrzelewyMozliwePozycjePobierzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

Public Class RozliczeniePrzelewPlikZatwierdzWynik
    Public dane As DataSet
    Public status_opis As String
    Public status As Integer
End Class

#End Region

#Region "ISKA"
Public Class AwizoIskaPozycjeWyslijWynik
    Public status As Integer
    Public status_opis As String
    Public awizoId As Integer
    Public user As String
    Public password As String
    Public url As String
    Public dane As DataSet
End Class
Public Class AwizoIskaPotwierdzenieZamknijIZapiszWynik
    Public status As Integer
    Public status_opis As String
End Class

#End Region

#Region "Dev test"
Public Class ProjektLogoDodajWynik
    Public status As Integer
    Public status_opis As String
End Class

Public Class SzablonWydrukuGrafikaDodajWynik
    Public status As Integer
    Public status_opis As String
End Class
#End Region

Public Enum ListaTypeEnum
    ID
    WARTOSC_INT
    WARTOSC_NVARCHAR
    WARTOSC_BIN
End Enum