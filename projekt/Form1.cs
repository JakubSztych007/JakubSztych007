




// Deklaracje bibliotek u�ywanych w projekcie.
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection.Metadata.Ecma335;
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using System.Reflection.Emit;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Diagrams;

// Definicja przestrzeni nazw projektu.
namespace projekt
{
    // Definicja g��wnej klasy Form1, dziedzicz�cej po klasie Form.
    public partial class Form1 : Form
    {

        // Metoda do czytania danych z pliku Excel.
        void readExcel()
        {
            // Pobranie �cie�ki do pliku i otwarcie pliku Excel.
            string FilePath = Convert.ToString(wyszukaj.FileName);
            using (var workbook = new XLWorkbook(FilePath))
            {
                var ws = workbook.Worksheet(1); // Otwarcie pierwszego arkusza.
                int i = 10; // Zmienna pomocnicza do iterowania po wierszach.
                int x = 0; // Zmienna pomocnicza do liczenia pustych kom�rek.

                // P�tla wykonuje si�, dop�ki dwie kolejne kom�rki w pierwszej kolumnie nie s� puste.
                while (!ws.Cell(i, 1).IsEmpty() || !ws.Cell(i - 1, 1).IsEmpty())
                {
                    x = 0;
                    // P�tla wykonuje si�, dop�ki kom�rka w drugiej kolumnie jest pusta.
                    while (ws.Cell(i, 2).IsEmpty() && x < 4)
                    {
                        i++;
                        x++;
                    }
                    // Je�eli kom�rka w drugiej kolumnie nie jest pusta, dodaj jej warto�� do comboBox1.
                    if (!ws.Cell(i, 2).IsEmpty())
                    {
                        comboBox1.Items.Add(ws.Cell(i, 2).Value.ToString());
                        i++;
                    }
                }
            }
        }
        OpenFileDialog wyszukaj = new OpenFileDialog();
        OpenFileDialog wyszukaj1 = new OpenFileDialog();
        OpenFileDialog wyszukaj2 = new OpenFileDialog();
        OpenFileDialog wyszukaj3 = new OpenFileDialog();
        // Konstruktor klasy Form1.
        public Form1()
        {
            InitializeComponent(); // Inicjalizacja komponent�w formularza.
            // Ustawienie warto�ci pocz�tkowych dla niekt�rych p�l tekstowych.
            textpropwyk.Text = 1.ToString();
            textpropcw.Text = 1.ToString();
            textproplab.Text = 1.ToString();
            textpropproj.Text = 1.ToString();
            textpropsem.Text = 1.ToString();

            // Ustawienie w�a�ciwo�ci dla comboBox1, aby uniemo�liwi� u�ytkownikowi wpisywanie warto�ci.
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            // Ustawienie filtra dla dialogu wyboru plik�w, aby akceptowa� tylko pliki programu Excel.
            wyszukaj.Filter = "Pliki programu Excel|*.xlsx";
            wyszukaj1.Filter = "Pliki programu Excel|*.xlsx";
            wyszukaj2.Filter = "Pliki programu Excel|*.xlsx";
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        // Metoda wywo�ywana po klikni�ciu przycisku button5.
        private void button5_Click(object sender, EventArgs e)
        {

            float.TryParse(textECTS.Text, out float ECTS);
            float.TryParse(textECTSNA.Text, out float ECTSNA);
            int liczbaNiezerowych = 0;
            int q = 0;
            int w = 0;
            // Tworzenie listy przechowuj�cej niepuste dane z TextBox-�w.
            List<string> data = new List<string>();

            if (textgodzinywyk.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w wyk�adach/ " + textgodzinywyk.Text); } //s
            if (textgodzinycw.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w �wiczeniach/ " + textgodzinycw.Text); } //s
            if (textgodzinylab.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w laboratoriach/ " + textgodzinylab.Text); } //s
            if (textgodzinyproj.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w godzinach projektowych/ " + textgodzinyproj.Text); } //s
            if (textgodzinysem.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w seminariach/ " + textgodzinysem.Text); } //s
            if (textnauproj.Text != "0") { q++; liczbaNiezerowych++; data.Add("Realizowanie projektu pod kierunkiem nauczyciela/ " + textnauproj.Text); } //s
            if (textgodzinykon.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w konsultacjach/ " + textgodzinykon.Text); } //s
            if (textegz.Text != "0") { q++; liczbaNiezerowych++; data.Add("Udzia� w egzaminie/ " + textegz.Text); } //s

            if (textsamwyk.Text != "0") { q++; data.Add("Samodzielna studiowanie wyk�ad�w/ " + textsamwyk.Text); }
            if (textsamcw.Text != "0") { q++; data.Add("Samodzielnie przygotowanie do �wicze�/ " + textsamcw.Text); }
            if (textsamlab.Text != "0") { q++; data.Add("Samodzielne przygotowanie do laboratori�w/ " + textsamlab.Text); }
            if (textsamproj.Text != "0") { q++; data.Add("Samodzielna realizacja projekt�w/ " + textsamproj.Text); }
            if (textsamsem.Text != "0") { q++; data.Add("Samodzielne przygotowanie do seminari�w/ " + textsamsem.Text); }
            if (textsamegz.Text != "0") { q++; data.Add("Przygotowanie do egzaminu/ " + textsamegz.Text); }
            if (textsamzal.Text != "0") { q++; data.Add("Przygotowanie do zaliczenia/ " + textsamzal.Text); }

            if (!textstuwyz.Text.Equals("0"))
            {
                w++;
                data.Add("Sumaryczne obci��enie prac� studenta (" + textstuwyz.Text);
            }
            if (!textnauwyz.Text.Equals("0"))
            {
                w++;
                data.Add("Zaj�cia z udzia�em nauczycieli (" + textnauwyz.Text);
            }

            string wynik = "";

            for (int i = 1; i <= liczbaNiezerowych; i++)
            {
                wynik += i.ToString();
                if (i < liczbaNiezerowych)
                {
                    wynik += "+";
                }
            }

            // Przygotowanie finalnego ci�gu do skopiowania.
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < q; i++)
            {
                stringBuilder.AppendLine((i + 1).ToString() + ". " + data[i]);
            }

            // Dodanie ostatnich dw�ch element�w bez numeracji, ale tylko je�li nie s� r�wne "0".
            if (w >= 1)
            {
                stringBuilder.AppendLine(data[q] + ") / " + ECTS);
            }
            if (w == 2)
            {
                stringBuilder.AppendLine(data[q + 1] + ") / (" + wynik + ") / " + ECTSNA);
            }


            // Kopiowanie danych do schowka systemowego, je�li jakiekolwiek dane s� niepuste.
            string copiedData = stringBuilder.ToString();
            if (!string.IsNullOrEmpty(copiedData))
            {
                Clipboard.SetText(copiedData);
                MessageBox.Show("Dane zosta�y skopiowane do schowka.");
            }
            else
            {
                MessageBox.Show("�adne dane nie zosta�y skopiowane (wszystkie TextBox-y maj� warto�� 0).");
            }
        }
        // Metoda sprawdzaj�ca, czy pole tekstowe nie jest puste. Je�li jest puste, ustawia jego warto�� na "0".
        private void EnsureTextBoxIsNotEmpty(TextBox textBox)
        {
            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "0";
            }
        }

        // Metoda wywo�ywana podczas �adowania formularza.
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Metody obs�uguj�ce zdarzenie zmiany tekstu w polach tekstowych. Obecnie puste.
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        // Metody obs�uguj�ce zmian� warto�ci pask�w przewijania.
        private void trackBarwyk_ValueChange(object sender, EventArgs e)
        {

        }

        private void textpropwyk_TextChanged(object sender, EventArgs e)
        {

        }

        private void textproppro_TextChanged(object sender, EventArgs e)
        {

        }

        // Metoda obs�uguj�ca zmian� warto�ci paska przewijania dla �wicze�. Oblicza warto�� i wy�wietla w odpowiednim polu tekstowym.
        private void trackBarcw_ValueChanged(object sender, EventArgs e)
        {
            float trueValue = trackBarcw.Value / 10f;
            textpropcw.Text = trueValue.ToString("0.0");
        }

        private void trackBarwyk_ValueChanged(object sender, EventArgs e)
        {
            float trueValue = trackBarwyk.Value / 10f;
            textpropwyk.Text = trueValue.ToString("0.0");
        }

        private void trackBarlab_ValueChanged(object sender, EventArgs e)
        {
            float trueValue = trackBarlab.Value / 10f;
            textproplab.Text = trueValue.ToString("0.0");
        }

        private void trackBarproj_ValueChanged(object sender, EventArgs e)
        {
            float trueValue = trackBarproj.Value / 10f;
            textpropproj.Text = trueValue.ToString("0.0");
        }

        private void trackBarsem_ValueChanged(object sender, EventArgs e)
        {
            float trueValue = trackBarsem.Value / 10f;
            textpropsem.Text = trueValue.ToString("0.0");
        }

        // Metody obs�uguj�ce zdarzenie zmiany tekstu w pozosta�ych polach tekstowych. Obecnie puste.
        private void textsamwyk_TextChanged(object sender, EventArgs e)
        {

        }

        private void textzajecia_TextChanged(object sender, EventArgs e)
        {

        }

        // Metoda wywo�ywana po klikni�ciu przycisku. Pobiera dane z pliku Excel i uzupe�nia pola tekstowe w formularzu.
        private void button1_Click_1(object sender, EventArgs e)
        {

            // Sprawd�, czy ComboBox 'comboBox1' ma wybran� jak�� warto��.
            if (comboBox1.SelectedItem == null || string.IsNullOrWhiteSpace(comboBox1.Text))
            {
                // Je�li nie ma wybranej warto�ci, wy�wietl komunikat.
                MessageBox.Show("Nie wybra�e� przedmiotu.", "Uwaga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {


                int index = comboBox1.SelectedIndex; // Pobranie wybranego indeksu z combobox'a.
                IXLWorksheet worksheet;
                string FilePath = Convert.ToString(wyszukaj.FileName); // Pobranie �cie�ki do pliku.
                using (var workbook = new XLWorkbook(FilePath)) // Otwarcie pliku Excel.
                {
                    worksheet = workbook.Worksheet(1); // Otwarcie pierwszego arkusza.
                    var cellValueA20 = worksheet.Cell("A20").Value.ToString(); // Pobranie warto�ci z kom�rki A20.

                    if (string.IsNullOrEmpty(cellValueA20)) // Sprawdzenie, czy kom�rka A20 jest pusta.
                    {
                        index += 100; // Zwi�kszenie indeksu o 100, je�li A20 jest pusta.
                    }
                    //MessageBox.Show(index + " tyle wynosi index");
                    int formNumber = Program.GetFormNumber(index); // Pobranie numeru formularza na podstawie indeksu.

                    // Logika wybieraj�ca odpowiednie kolumny w zale�no�ci od wybranego indeksu.
                    string x = "";
                    string y = "";
                    string z = "";

                    // Tutaj znajduj� si� instrukcje warunkowe ustawiaj�ce zmienne x, y, z na podstawie warto�ci indeksu.
                    if ((index >= 0 && index <= 4 || index >= 9 && index <= 14) || (index >= 100 && index <= 104 || index >= 108 && index <= 113))
                    {
                        x = "P";
                        y = "R";
                        z = "Q";

                    }
                    else if ((index >= 5 && index <= 6) || (index == 105 || index == 107))
                    {
                        x = "E";
                        y = "F";
                        z = "T";
                    }
                    else if (index == 7 || index == 106)
                    {
                        x = "S";
                        y = "U";
                        z = "Z";
                    }
                    else if ((index >= 15 && index <= 21 || index == 8) || (index >= 114 && index <= 120))
                    {
                        x = "S";
                        y = "U";
                        z = "T";
                    }
                    else if ((index >= 22 && index <= 29) || (index >= 121 && index <= 128))
                    {
                        x = "V";
                        y = "X";
                        z = "W";
                    }
                    else if ((index >= 30 && index <= 36) || (index >= 129 && index <= 135))
                    {
                        x = "Y";
                        y = "AA";
                        z = "Z";
                    }
                    else if ((index >= 37 && index <= 44 || index >= 52 && index <= 56) || (index >= 136 && index <= 143 || index >= 151 && index <= 155))
                    {
                        x = "AB";
                        y = "AD";
                        z = "AC";
                    }
                    else if ((index >= 45 && index <= 50 || index >= 57 && index <= 62) || (index >= 144 && index <= 149 || index >= 156 && index <= 161))
                    {
                        x = "AE";
                        y = "AG";
                        z = "AF";
                    }
                    else if (index == 51 || index == 63 || index == 150 || index == 162)
                    {
                        x = "AH";
                        y = "AJ";
                        z = "AI";
                    }
                    else
                    {
                        x = "O";
                        y = "AJ";
                        z = "AI";
                    }



                    // Pobieranie warto�ci z arkusza i ustawianie ich w odpowiednich polach tekstowych.
                    var value1 = worksheet.Cell(x + formNumber).Value.ToString();
                    var value2 = worksheet.Cell("K" + formNumber).Value.ToString();
                    var value3 = worksheet.Cell("L" + formNumber).Value.ToString();
                    var value4 = worksheet.Cell("M" + formNumber).Value.ToString();
                    var value5 = worksheet.Cell("N" + formNumber).Value.ToString();
                    var value6 = worksheet.Cell("O" + formNumber).Value.ToString();
                    var value7 = worksheet.Cell(y + formNumber).Value.ToString();
                    var value8 = worksheet.Cell("J" + formNumber).Value.ToString();
                    var value9 = worksheet.Cell(z + formNumber).Value.ToString();
                    var value10 = worksheet.Cell("B" + formNumber).Value.ToString();



                    textgodzinyall.Text = value1;
                    textgodzinywyk.Text = value2;
                    textgodzinycw.Text = value3;
                    textgodzinylab.Text = value4;
                    textgodzinyproj.Text = value5;
                    textgodzinysem.Text = value6;
                    textECTS.Text = value7;
                    textECTSNA.Text = value8;
                    label1.Text = value10;

                    if (value9 == "E")
                    {
                        int textEgz = 2;
                        int textPe = 15;
                        textegz.Text = textEgz.ToString();
                        textsamegz.Text = textPe.ToString();
                        int textPz = 0;
                        textsamzal.Text = textPz.ToString();
                    }
                    else
                    {
                        int textPz = 15;
                        int textPe = 0;
                        textsamegz.Text = textPe.ToString();
                        textsamzal.Text = textPz.ToString();
                    }
                    if (label1.Text == "Praktyka zawodowa - 4 tygodnie")
                    {
                        textgodzinywyk.Text = "";
                    }

                    // Wywo�anie metody EnsureTextBoxIsNotEmpty dla ka�dego pola tekstowego, aby upewni� si�, �e �adne z nich nie jest puste.
                    EnsureTextBoxIsNotEmpty(textgodzinyall);
                    EnsureTextBoxIsNotEmpty(textgodzinywyk);
                    EnsureTextBoxIsNotEmpty(textgodzinycw);
                    EnsureTextBoxIsNotEmpty(textgodzinylab);
                    EnsureTextBoxIsNotEmpty(textgodzinyproj);
                    EnsureTextBoxIsNotEmpty(textgodzinysem);
                    EnsureTextBoxIsNotEmpty(textECTS);
                    EnsureTextBoxIsNotEmpty(textECTSNA);
                    EnsureTextBoxIsNotEmpty(textegz);

                    // Parsowanie tekst�w na warto�ci liczbowe, obliczenia i aktualizacja odpowiednich p�l tekstowych.
                    double godzinyAll = double.Parse(textgodzinyall.Text);
                    double godzinyWyk = double.Parse(textgodzinywyk.Text);
                    double godzinyCw = double.Parse(textgodzinycw.Text);
                    double godzinyLab = double.Parse(textgodzinylab.Text);
                    double godzinyProj = double.Parse(textgodzinyproj.Text);
                    double godzinySem = double.Parse(textgodzinysem.Text);
                    double ECTS = double.Parse(textECTS.Text);
                    double ECTSNA = double.Parse(textECTSNA.Text);

                    double textgodzinyKon = godzinyAll - (godzinyWyk + godzinyCw + godzinyLab + godzinyProj + godzinySem);

                    textgodzinykon.Text = textgodzinyKon.ToString();
                }
            }
        }
        // Metoda wywo�ywana przez timer. Wykonuje obliczenia na podstawie danych wprowadzonych przez u�ytkownika i aktualizuje interfejs.
        private void timer_Tick(object sender, EventArgs e)
        {
            // Parsowanie tekst�w na warto�ci liczbowe, obliczenia i aktualizacja odpowiednich p�l tekstowych.
            double.TryParse(textgodzinywyk.Text, out double godzinyWyk);
            double.TryParse(textgodzinycw.Text, out double godzinyCw);
            double.TryParse(textgodzinylab.Text, out double godzinyLab);
            double.TryParse(textgodzinyproj.Text, out double godzinyProj);
            double.TryParse(textgodzinysem.Text, out double godzinySem);
            double.TryParse(textECTS.Text, out double ECTS);
            double.TryParse(textECTSNA.Text, out double ECTSNA);
            double.TryParse(textpropwyk.Text, out double textpropWyk);
            double.TryParse(textpropcw.Text, out double textpropCw);
            double.TryParse(textproplab.Text, out double textpropLab);
            double.TryParse(textpropproj.Text, out double textpropProj);
            double.TryParse(textpropsem.Text, out double textpropSem);
            double.TryParse(textgodzinykon.Text, out double textgodzinyKon);
            double.TryParse(textegz.Text, out double textEgz);
            double.TryParse(textsamegz.Text, out double textsamEgz);
            double.TryParse(textsamzal.Text, out double textsamZal);
            double.TryParse(textnauproj.Text, out double textnauProj);

            double resultw = godzinyWyk * textpropWyk;
            double textsamWyk = Math.Round(resultw);
            textsamwyk.Text = textsamWyk.ToString();

            double resultc = godzinyCw * textpropCw;
            double textsamCw = Math.Round(resultc);
            textsamcw.Text = textsamCw.ToString();

            double resultl = godzinyLab * textpropLab;
            double textsamLab = Math.Round(resultl);
            textsamlab.Text = textsamLab.ToString();

            double resultp = godzinyProj * textpropProj;
            double textsamProj = Math.Round(resultp);
            textsamproj.Text = textsamProj.ToString();

            double results = godzinySem * textpropSem;
            double textsamSem = Math.Round(results);
            textsamsem.Text = textsamSem.ToString();

            double resultnwyz = godzinyWyk + godzinyCw + godzinyLab + godzinySem + godzinyProj + textgodzinyKon + textEgz + textnauProj;
            double textnauWyz = Math.Round(resultnwyz);
            textnauwyz.Text = textnauWyz.ToString();

            double resultswyz = textsamWyk + textsamCw + textsamLab + textsamProj + textsamSem + textgodzinyKon + textEgz + textsamEgz + textsamZal + godzinyWyk + godzinyCw + godzinyLab + godzinySem + godzinyProj;
            double textstuWyz = Math.Round(resultswyz);
            textstuwyz.Text = textstuWyz.ToString();

            double resultswyn = ECTS * 30;
            double textstuWyn = Math.Round(resultswyn);
            textstuwyn.Text = textstuWyn.ToString();

            double resultnwyn = ECTSNA * 30;
            double textnauWyn = Math.Round(resultnwyn);
            textnauwyn.Text = textnauWyn.ToString();

            double resultsmin = (ECTS - 0.25) * 30;
            double textstuMin = Math.Round(resultsmin);
            textstumin.Text = textstuMin.ToString();

            double resultsmaks = (ECTS + 0.25) * 30;
            double textstuMaks = Math.Round(resultsmaks);
            textstumaks.Text = textstuMaks.ToString();

            double resultnmin = (ECTSNA - 0.25) * 30;
            double textnauMin = Math.Round(resultnmin);
            textnaumin.Text = textnauMin.ToString();

            double resultnmaks = (ECTSNA + 0.25) * 30;
            double textnauMaks = Math.Round(resultnmaks);
            textnaumaks.Text = textnauMaks.ToString();

            // Zmiana koloru t�a pola tekstowego na zielony lub czerwony w zale�no�ci od spe�nienia warunk�w.
            double valueMin = double.Parse(textstumin.Text);
            double valueMax = double.Parse(textstumaks.Text);
            double valuE = double.Parse(textstuwyz.Text);

            if (valuE >= valueMin && valuE <= valueMax)
            {
                textstuwyz.BackColor = System.Drawing.Color.Green;

            }
            else
            {
                textstuwyz.BackColor = System.Drawing.Color.Red;

            }


            double valueMinn = double.Parse(textnaumin.Text);
            double valueMaxn = double.Parse(textnaumaks.Text);
            double valuEn = double.Parse(textnauwyz.Text);

            if (valuEn >= valueMinn && valuEn <= valueMaxn)
            {
                textnauwyz.BackColor = System.Drawing.Color.Green;

            }
            else
            {
                textnauwyz.BackColor = System.Drawing.Color.Red;

            }
        }

        // Metody obs�uguj�ce zmian� tekstu w polach tekstowych. Obecnie puste.
        private void textECTSNA_TextChanged(object sender, EventArgs e)
        {

        }

        public void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // Metoda wywo�ywana po klikni�ciu przycisku, zamyka aplikacj�.
        private void button1_Click_3(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // Metoda wywo�ywana po klikni�ciu przycisku, otwiera dialog wyboru pliku.
        private void button2_Click(object sender, EventArgs e)
        {
            if (wyszukaj.ShowDialog() == DialogResult.OK)
            { }
        }

        // Metoda wywo�ywana po klikni�ciu przycisku, sprawdza czy plik zosta� wybrany i wywo�uje metod� readExcel().
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(wyszukaj.FileName))
            {
                MessageBox.Show("Poprawnie wybrano plik");
            }
            else
            {
                MessageBox.Show("Nie wybrano pliku");
                return;
            }
            readExcel();
        }

        // Metody obs�uguj�ce zmian� tekstu w polach tekstowych. Obecnie puste.
        private void textstuwyz_TextChanged(object sender, EventArgs e)
        {

        }

        private void textnauwyz_TextChanged(object sender, EventArgs e)
        {

        }

        // Metoda wywo�ywana po klikni�ciu przycisku, wykonuje obliczenia na podstawie danych wprowadzonych przez u�ytkownika.
        private void button4_Click(object sender, EventArgs e)
        {
            // Parsowanie tekst�w na warto�ci liczbowe, obliczenia i aktualizacja odpowiedniego pola tekstowego.
            int.TryParse(textgodzinyall.Text, out int godzinyAll);
            float.TryParse(textECTSNA.Text, out float ECTSNA);


            float textgodzinyKon = (ECTSNA * 30) - godzinyAll;
            textgodzinykon.Text = textgodzinyKon.ToString();
        }


        // Metoda obs�uguj�ca zdarzenie klikni�cia przycisku o identyfikatorze 8.
        private void button8_Click(object sender, EventArgs e)
        {
            // Clear all textboxes before starting the operation
            for (int i = 1; i <= 40; i++) // Assuming you have 40 textboxes as per your existing code logic
            {
                var textBox = this.Controls.Find("textBox" + i, true).FirstOrDefault() as TextBox;
                if (textBox != null)
                {
                    textBox.Text = string.Empty;
                }
            }
            bool dataFound = false;

            if (label1.Text != "Nazwa przedmiotu")
            {
                string FilePath1 = Convert.ToString(wyszukaj1.FileName);
                if (File.Exists(FilePath1))
                {
                    using (var workbook1 = new XLWorkbook(FilePath1))
                    {
                        var ws1 = workbook1.Worksheet(1);

                        var range = ws1.RangeUsed();
                        foreach (var row in range.RowsUsed())
                        {
                            if (row.Cell(1).Value.ToString() == label1.Text)
                            {
                                dataFound = true;
                                for (int i = 1; i <= 20; i++)
                                {
                                    var textBox = this.Controls.Find("textBox" + i, true).FirstOrDefault() as TextBox;
                                    if (textBox != null)
                                    {
                                        textBox.Text = row.Cell(i + 1).Value.ToString();
                                    }
                                }
                                break;
                            }
                        }
                    }

                    if (!dataFound)
                    {
                        MessageBox.Show("Nie znaleziono danych dla wybranego przedmiotu.");
                    }
                }
                else
                {
                    MessageBox.Show("Plik wskazuj�cy kody efekt�w nie zosta� wybrany.");
                }
            }
            else
            {
                MessageBox.Show("Nie wybra�e� przedmiotu");
            }

            // U�ycie drugiej lokalizacji do wyszukiwania pliku
            string FilePath2 = Convert.ToString(wyszukaj2.FileName);
            if (File.Exists(FilePath2))
            {
                using (var workbook2 = new XLWorkbook(FilePath2))
                {
                    var ws2 = workbook2.Worksheet(1);

                    for (int i = 1; i <= 20; i++)
                    {
                        var searchText = this.Controls.Find("textBox" + i, true).FirstOrDefault() as TextBox;
                        if (searchText != null && !string.IsNullOrEmpty(searchText.Text))
                        {
                            var matchedRow = ws2.RowsUsed().FirstOrDefault(r => r.Cell(1).Value.ToString().Equals(searchText.Text, StringComparison.OrdinalIgnoreCase));
                            if (matchedRow != null)
                            {
                                var value = matchedRow.Cell(2).Value;
                                var targetTextBox = this.Controls.Find("textBox" + (i + 20), true).FirstOrDefault() as TextBox;
                                if (targetTextBox != null)
                                {
                                    targetTextBox.Text = value.ToString();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Plik wskazuj�cy opisy efekt�w nie zosta� wybrany.");
            }
        }

        // Metoda obs�uguj�ca zdarzenie klikni�cia przycisku o identyfikatorze 9.
        private void button9_Click(object sender, EventArgs e)
        {

            // Sprawdzenie, czy textBox1 nie jest pusty i skopiowanie jego zawarto�ci do schowka.
            if (!string.IsNullOrEmpty(textBox1.Text))
            {

                var stringBuilder = new System.Text.StringBuilder();

                int wCounter = 1;
                int uCounter = 1;
                int kCounter = 1;

                for (int i = 1; i <= 20; i++)
                {
                    var sourceTextBox = this.Controls.Find("textBox" + i, true).FirstOrDefault() as TextBox;
                    var targetTextBox = this.Controls.Find("textBox" + (i + 20), true).FirstOrDefault() as TextBox;

                    if (sourceTextBox != null && targetTextBox != null && !string.IsNullOrEmpty(sourceTextBox.Text))
                    {
                        string prefix;
                        if (sourceTextBox.Text.ToUpper().Contains("W"))
                        {
                            prefix = $"W{wCounter++}_";
                        }
                        else if (sourceTextBox.Text.ToUpper().Contains("U"))
                        {
                            prefix = $"U{uCounter++}_";
                        }
                        else
                        {
                            prefix = $"K{kCounter++}_";
                        }

                        stringBuilder.AppendLine(prefix + targetTextBox.Text + "/" + sourceTextBox.Text);
                    }
                }

                string resultString = stringBuilder.ToString();
                if (!string.IsNullOrEmpty(resultString))
                {
                    Clipboard.SetText(resultString);
                    MessageBox.Show("Efekty zosta�y skopiowane");
                }
            }
            else
            {
                MessageBox.Show("Brak tekstu do skopiowania.");
            }

        }

        // Metoda obs�uguj�ca zdarzenie klikni�cia przycisku o identyfikatorze 10.
        private void button10_Click(object sender, EventArgs e)
        {
            textopis.Text = string.Empty;
            string FilePath = Convert.ToString(wyszukaj3.FileName);
            // Sprawdzenie, czy plik Excel istnieje.
            if (!File.Exists(FilePath))
            {
                MessageBox.Show("Plik do pobrania opisu przedmiotu nie zosta� wybrany");
                return;
            }

            using (var workbook = new XLWorkbook(FilePath))
            {
                var worksheet = workbook.Worksheet(1); // Praca z pierwszym arkuszem.

                bool found = false; // Zmienna do �ledzenia, czy znaleziono szukany tekst.
                foreach (var row in worksheet.RangeUsed().Rows())
                {
                    // Sprawd�, czy w pierwszej kolumnie jest szukany tekst
                    if (row.Cell(1).Value.ToString().Equals(label1.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        // Za��my, �e dane do textopis s� w drugiej kolumnie
                        textopis.Text = row.Cell(2).Value.ToString();
                        found = true;
                        break; // Przerwij p�tl� po znalezieniu odpowiedniego wiersza
                    }
                }

                if (!found)
                {
                    MessageBox.Show("Nie znaleziono opisu dla tego przedmiotu");
                }
            }
        }

        // Metoda obs�uguj�ca zdarzenie zmiany tekstu w textopis.
        private void textopis_TextChanged(object sender, EventArgs e)
        {

        }

        // Metoda obs�uguj�ca zdarzenie klikni�cia przycisku o identyfikatorze 11.
        private void button11_Click(object sender, EventArgs e)
        {
            // Sprawdzenie, czy textopis nie jest pusty i skopiowanie jego zawarto�ci do schowka.
            if (!string.IsNullOrEmpty(textopis.Text))
            {
                Clipboard.SetText(textopis.Text);
                MessageBox.Show("Opis zosta� skopiowany");
            }
            else
            {
                MessageBox.Show("Brak tekstu do skopiowania.");
            }
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }


        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (wyszukaj1.ShowDialog() == DialogResult.OK)
            {
                // Kod do obs�ugi po wybraniu pliku przez wyszukaj1
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (wyszukaj2.ShowDialog() == DialogResult.OK)
            {
                // Kod do obs�ugi po wybraniu pliku przez wyszukaj2
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (wyszukaj3.ShowDialog() == DialogResult.OK)
            {
                // Kod do obs�ugi po wybraniu pliku przez wyszukaj3
            }
        }

        private void textnaumin_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            trackBarwyk.Value = 10;
            textpropwyk.Text = "1";
            trackBarcw.Value = 10;
            textpropcw.Text = "1";
            trackBarlab.Value = 10;
            textproplab.Text = "1";
            trackBarproj.Value = 10;
            textpropproj.Text = "1";
            trackBarsem.Value = 10;
            textpropsem.Text = "1";
        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }
    }
}