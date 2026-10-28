using System;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;
using cdn_api;
using System.Collections.Generic;
using System.Linq;
using ONE034;

namespace one034
{
    public partial class ONE034Form : Form
    {
        
        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.137.149";
        private string XLSQLDatabase = "PoloMarket";//"ERPXL_PoloMarket"
        private string XLUser = "admin"; //
        private string XLPass = "NE!w0rk12";//"Dupa2003!";
        string connectionString = @"Data Source = SRV1; Initial Catalog = ERPXL_PoloMarket; User ID = admin; Password=Dupa2003!";
        string date1 = "1800-12-28";
        
        //@"Data Source = SQLSERVER; Initial Catalog = DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";

        public ONE034Form()
        {
            //this.connectionString = GetConnectionString();
            InitializeComponent();
        }


        void Button1_Click(object sender, EventArgs e)
        {
            #region excel5
                #region wybranie pliku
                Cursor.Current = Cursors.WaitCursor;
                if (tB1.Text == "nie wybrano pliku")
                {
                    MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string filePath = tB1.Text;
                DataTable dataTable;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false//true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    dataTable = dataSet.Tables[0];
                }
            #endregion

            #region Przypisywanie danych z datasetu do listy przyszłych pozycji.

            SprawdzamCzyArkuszNieJestPusty(dataTable);
                dodajDokumenty(dataTable);

                #endregion

            #endregion


        }
        #region Dodawanie dokumetów
        void dodajDokumenty(DataTable dataTable)
        {


            XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
            nowyUzytkownik.Wersja = XLWersjaAPI;
            nowyUzytkownik.Winieta = XLWinieta;
            nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
            nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
            nowyUzytkownik.SerwerKlucza = XLSQLServer;
            nowyUzytkownik.Baza = XLSQLDatabase;
            nowyUzytkownik.OpeIdent = XLUser;
            nowyUzytkownik.OpeHaslo = XLPass;
            nowyUzytkownik.PortHTTPSerweraKlucza = 1;
            nowyUzytkownik.RodzajSerweraKlucza = 0;
            nowyUzytkownik.SerwerKlucza = "SRV1::5000139268";
            nowyUzytkownik.ProgramID = "ONES01 - Dodawanie Faktury";

            cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);
            
            int i = 1;

            
            string akronim = Convert.ToString(dataTable.Rows[i][3]).ToUpper();
            int count = 0;
            int count2 = 0;


            for (i = 1; i < dataTable.Rows.Count; i++)
            {
                akronim = Convert.ToString(dataTable.Rows[i][3]).ToUpper();
                string query6 = "SELECT COUNT(*) FROM [ERPXL_PoloMarket].[CDN].[KntGrupy] WHERE KnG_Akronim = '" + akronim + "';";
                if (Convert.ToInt32(ConnectToSQL(connectionString, query6)) == 1)
                    count2++;
            }

            for (i = 1; i < dataTable.Rows.Count; i++)
            {
                string kod = Convert.ToString(dataTable.Rows[i][5]);
                string query7 = "SELECT COUNT(*) FROM  [ERPXL_PoloMarket].[CDN].[TwrGrupy] WHERE TwG_Kod='" + kod + "';";
                if (Convert.ToInt32(ConnectToSQL(connectionString, query7)) == 1)
                    count++;
            }

            if (count2== dataTable.Rows.Count - 1 && count == dataTable.Rows.Count-1)
            {
                List<string> akronimy = new List<string>();
                
                for (i = 1; i < dataTable.Rows.Count; i++)
                {
                    akronim = Convert.ToString(dataTable.Rows[i][3]).ToUpper();
                    if (!akronimy.Contains(akronim))
                        akronimy.Add(akronim);
                 
                }


                   pBar1.Visible = true;
                 // Set Minimum to 1 to represent the first file being copied.
                 pBar1.Minimum = 1;
                 // Set Maximum to the total number of files to copy.
                 pBar1.Maximum = akronimy.Count+1;
                 // Set the initial value of the ProgressBar.
                 pBar1.Value = 1;
                 // Set the Step property to a value of 1 to represent each file being copied.
                 pBar1.Step = 1;




                    
                foreach (string aAkronimy in akronimy)
                {
                    pBar1.PerformStep();

                    for (i = 1; i < dataTable.Rows.Count; i++)
                        if (Convert.ToString(dataTable.Rows[i][3]).ToUpper() == aAkronimy)
                            break;

                    

                    akronim = Convert.ToString(dataTable.Rows[i][3]).ToUpper();
                    string date2 = Convert.ToString(dataTable.Rows[i][0]) + "-" + Convert.ToString(dataTable.Rows[i][6]).Substring(5, 2) + "-" + Convert.ToString(dataTable.Rows[i][6]).Substring(8, 2);
                    string query = "SELECT KnG_GIDFirma FROM  [ERPXL_PoloMarket].[CDN].[KntGrupy] WHERE KnG_Akronim='" + akronim + "';";
                    string query2 = "SELECT KnG_GIDLp FROM  [ERPXL_PoloMarket].[CDN].[KntGrupy] WHERE KnG_Akronim='" + akronim + "';";
                    string query3 = "SELECT KnG_GIDNumer FROM  [ERPXL_PoloMarket].[CDN].[KntGrupy] WHERE KnG_Akronim='" + akronim + "';";
                    string query4 = "SELECT KnG_GIDTyp FROM  [ERPXL_PoloMarket].[CDN].[KntGrupy] WHERE KnG_Akronim='" + akronim + "';";
                    string query5 = "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); ";

                    XLDokumentNagInfo_20202 nowyDokument = new XLDokumentNagInfo_20202();
                    nowyDokument.Wersja = XLWersjaAPI;
                    nowyDokument.GIDFirma = 0;
                    nowyDokument.GIDLp = 0;
                    nowyDokument.GIDNumer = 0;
                    nowyDokument.GIDTyp = 0;
                    nowyDokument.KntFirma = Convert.ToInt32(ConnectToSQL(connectionString, query));
                    nowyDokument.KntLp = Convert.ToInt32(ConnectToSQL(connectionString, query2));
                    nowyDokument.KntNumer = Convert.ToInt32(ConnectToSQL(connectionString, query3));
                    nowyDokument.KntTyp = Convert.ToInt32(ConnectToSQL(connectionString, query4));
                    nowyDokument.Typ = 3; //FS
                    nowyDokument.Tryb = 2;
                    nowyDokument.Korekta = 0;
                    nowyDokument.Rok = Convert.ToInt32(dataTable.Rows[i][0]);
                    nowyDokument.Miesiac = Convert.ToInt32(dataTable.Rows[i][1]);
                    nowyDokument.IDOddzial = Convert.ToInt32(dataTable.Rows[i][2]);
                    nowyDokument.DataSpr = Convert.ToInt32(ConnectToSQL(connectionString, query5));
                    int dokumentId = 0;
                    cdn_api.cdn_api.XLNowyDokument(sesja, ref dokumentId, nowyDokument);



                    for (int k = 1; k < dataTable.Rows.Count; k++)
                    {
                        if (Convert.ToString(dataTable.Rows[k][3]).ToUpper() == aAkronimy)
                            {
                            

                            XLDokumentElemInfo_20202 pozycja = new XLDokumentElemInfo_20202();
                            pozycja.Wersja = XLWersjaAPI;
                            pozycja.Ilosc = "1";
                            pozycja.TowarKod = Convert.ToString(dataTable.Rows[k][5]);//"2010283773117";
                            pozycja.TowarEAN = "";
                            pozycja.TowarNazwa = Convert.ToString(dataTable.Rows[k][11]);
                            pozycja.Cena = Convert.ToString(dataTable.Rows[k][12]);
                            pozycja.Vat = Convert.ToString(dataTable.Rows[k][13]);
                            //pozycja.TypKorekty = 5;
                            //pozycja.GIDLpOrg = 0;
                            cdn_api.cdn_api.XLDodajPozycje(dokumentId, pozycja);

                            XLAtrybutInfo_20202 atrybut = new XLAtrybutInfo_20202();
                            atrybut.Wersja = XLWersjaAPI;
                            atrybut.GIDFirma = pozycja.GIDFirma;
                            atrybut.GIDLp = pozycja.GIDLp;
                            atrybut.GIDNumer = pozycja.GIDNumer;
                            atrybut.GIDSubLp = 0;
                            atrybut.GIDTyp = pozycja.GIDTyp;
                            atrybut.Wartosc = Convert.ToString(dataTable.Rows[k][8]);
                            atrybut.Klasa = "Pesel";
                            cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut);


                            XLAtrybutInfo_20202 atrybut2 = new XLAtrybutInfo_20202();
                            atrybut2.Wersja = XLWersjaAPI;
                            atrybut2.GIDFirma = pozycja.GIDFirma;
                            atrybut2.GIDLp = pozycja.GIDLp;
                            atrybut2.GIDNumer = pozycja.GIDNumer;
                            atrybut2.GIDSubLp = 0;
                            atrybut2.GIDTyp = pozycja.GIDTyp;
                            string date3 = Convert.ToString(dataTable.Rows[k][0]) + "-" + Convert.ToString(dataTable.Rows[k][6]).Substring(5, 2) + "-" + Convert.ToString(dataTable.Rows[k][6]).Substring(8, 2);
                            string Datequery = "SELECT DATEDIFF(day, '" + date1 + "','" + date3 + "'); ";
                            atrybut2.Wartosc = Convert.ToString(ConnectToSQL(connectionString, Datequery));
                            atrybut2.Klasa = "Data badania";
                            cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut2);

                            XLAtrybutInfo_20202 atrybut3 = new XLAtrybutInfo_20202();
                            atrybut3.Wersja = XLWersjaAPI;
                            atrybut3.GIDFirma = pozycja.GIDFirma;
                            atrybut3.GIDLp = pozycja.GIDLp;
                            atrybut3.GIDNumer = pozycja.GIDNumer;
                            atrybut3.GIDSubLp = 0;
                            atrybut3.GIDTyp = pozycja.GIDTyp;
                            //atrybut.AtrFirma = 1034247;
                            //atrybut.AtrLp = 0;
                            //atrybut.AtrNumer = 45;
                            // atrybut.AtrSublp = 0;
                            //atrybut.AtrTyp = 128
                            atrybut3.Wartosc = Convert.ToString(dataTable.Rows[k][7]);
                            atrybut3.Klasa = "Pacjent";
                            cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut3);


                            XLAtrybutInfo_20202 atrybut4 = new XLAtrybutInfo_20202();
                            atrybut4.Wersja = XLWersjaAPI;
                            atrybut4.GIDFirma = pozycja.GIDFirma;
                            atrybut4.GIDLp = pozycja.GIDLp;
                            atrybut4.GIDNumer = pozycja.GIDNumer;
                            atrybut4.GIDSubLp = 0;
                            atrybut4.GIDTyp = pozycja.GIDTyp;
                            atrybut4.Wartosc = Convert.ToString(dataTable.Rows[k][9]);
                            atrybut4.Klasa = "Wiek";
                            cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut4);

                            XLAtrybutInfo_20202 atrybut5 = new XLAtrybutInfo_20202();
                            atrybut5.Wersja = XLWersjaAPI;
                            atrybut5.GIDFirma = pozycja.GIDFirma;
                            atrybut5.GIDLp = pozycja.GIDLp;
                            atrybut5.GIDNumer = pozycja.GIDNumer;
                            atrybut5.GIDSubLp = 0;
                            atrybut5.GIDTyp = pozycja.GIDTyp;
                            atrybut5.Wartosc = Convert.ToString(dataTable.Rows[k][10]);
                            atrybut5.Klasa = "Lekarz kierujący";
                            cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut5);
                        }

                    }




                    XLZamkniecieDokumentuInfo_20202 zamkniecie = new XLZamkniecieDokumentuInfo_20202();
                    zamkniecie.Wersja = XLWersjaAPI;
                    zamkniecie.Tryb = 0;
                    cdn_api.cdn_api.XLZamknijDokument(dokumentId, zamkniecie);


                  /*  MessageBox.Show("Dodano fakturę sprzedaży dla "+aAkronimy, "Komunikat",
                         MessageBoxButtons.OK, MessageBoxIcon.Information);*/
                }
                MessageBox.Show("Dodano faktury sprzedaży", "Komunikat",
                     MessageBoxButtons.OK, MessageBoxIcon.Information);
                pBar1.Value = 1;
                pBar1.Visible = false;
            }
            else
            {
                MessageBox.Show("Dane w arkuszu są niezgodne z formatem", "Komunikat",
                     MessageBoxButtons.OK, MessageBoxIcon.Warning);
                pBar1.Value = 1;
                pBar1.Visible = false;
            }

            cdn_api.cdn_api.XLLogout(sesja);


        }


        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        #endregion

        #region funkcje wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion

        #region funkcje pobranie danych z pliku
        bool SprawdzamCzyArkuszNieJestPusty(DataTable _dataTable)
        {
            if (_dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Wybrany arkusz pusty.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return true;
            }
            return false;
        }

        private void title1_Click(object sender, EventArgs e)
        {

        }

        private void ONE034Form_Load(object sender, EventArgs e)
        {

        }

        /*
        public bool CzyTraNag(short typ)        //zwraca informacje, czy zaznaczony dokument będzie się znajdował w tabeli TraNag
        {
            if (typ == 1529
            || typ == 434
            || typ == 800
            || typ == 1521
            || typ == 1529
            || typ == 1824
            || typ == 1828
            || typ == 1832
            || typ == 1836
            || typ == 2033
            || typ == 2034
            || typ == 2037
            || typ == 2041
            || typ == 2042
            || typ == 2045
            || typ == 2768
            || typ == 2832
            || typ == 2976
            || typ == 2977
            || typ == 3344
            || typ == 3352
            || typ == 4145
            || typ == 7680
            || typ == 7684)
                return true;
            else return false;
        }*/


    }
}
#endregion