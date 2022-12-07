using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace gyak12_excel
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB;     // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

        public void CreateTable()
        {

            string[] fejlécek = new string[] {
        "Kérdés",
        "1. válasz",
        "2. válaszl",
        "3. válasz",
        "Helyes válasz",
        "kép"};

            for (int i = 1; i < fejlécek.Length; i++)
            {
                xlSheet.Cells[1, 1] = fejlécek[0];
                xlSheet.Cells[1, 2] = fejlécek[1];
                xlSheet.Cells[1, 3] = fejlécek[2];
                xlSheet.Cells[1, 4] = fejlécek[3];
                xlSheet.Cells[1, 5] = fejlécek[4];
                xlSheet.Cells[1, 6] = fejlécek[5];
            }

            Models.HajosContext hajosContext = new Models.HajosContext();
            var mindenKérdés = hajosContext.Questions.ToList();

            object[,] adatTömb = new object[mindenKérdés.Count(), fejlécek.Count()];


            int sorokSzáma = adatTömb.GetLength(0);
            int oszlopokSzáma = adatTömb.GetLength(1);

            for (int i = 0; i < mindenKérdés.Count(); i++)
            {
                adatTömb[i, 0] = mindenKérdés[i].Question1;
                adatTömb[i, 1] = mindenKérdés[i].Answer1;
                adatTömb[i, 2] = mindenKérdés[i].Answer2;
                adatTömb[i, 3] = mindenKérdés[i].Answer3;
                adatTömb[i, 4] = mindenKérdés[i].CorrectAnswer;
                adatTömb[i, 5] = mindenKérdés[i].Image;
            }

            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSzáma, oszlopokSzáma);
            adatRange.Value2 = adatTömb;
            adatRange.Columns.AutoFit();

            Excel.Range fejllécRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejllécRange.Font.Bold = true;
            fejllécRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejllécRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejllécRange.EntireColumn.AutoFit();
            fejllécRange.RowHeight = 40;
            fejllécRange.Interior.Color = Color.Fuchsia;
            fejllécRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            adatRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range ElsoOszlop = xlSheet.get_Range("A1", Type.Missing).get_Resize(sorokSzáma, 1);
            ElsoOszlop.Font.Bold = true;

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range UtolosOszlop = xlSheet.get_Range("F2", Type.Missing).get_Resize(lastRowID , 1);
            UtolosOszlop.Interior.Color = Color.LightGreen;
            //nem találtam meg a kerekítõs részt googleon
        }

        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következõ feladatrészben következik

                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }


    }
}