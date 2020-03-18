using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace SchoolSort
{
    public partial class MainWindow : Window
    {
        private struct School
        {
            public string name;
            public int plan;
            public int two;
            public int three;
            public int four;
            public int five;
        }

        private const int smechenie1 = 3;
        private const int smechenie2 = 7;

        private List<School> SchoolMas = new List<School>();
        private string pathIn;
        private string pathOut;
        public MainWindow()
        {
            InitializeComponent();
        }
        private async void Calculate_Click(object sender, RoutedEventArgs e)
        {
            FilePathBut.IsEnabled = false;
            ExitBut.IsEnabled = false;
            CalculateBut.IsEnabled = false;
            await Task.Run(() => { Calculate(); });
            await Task.Run(() => { Outpute(); });
            FilePathBut.IsEnabled = true;
            ExitBut.IsEnabled = true;
            CalculateBut.IsEnabled = true;

        }
        private void Calculate()
        {
            Excel.Application ex = new Excel.Application();
            ex.Workbooks.Open(pathIn);

            string NexYacheika = (ex.Cells[smechenie1, smechenie1] as Excel.Range).Value2.ToString();
            string PrevYacheika = NexYacheika;
            School buf = new School
            {
                name = NexYacheika,
                plan = 1,
                two = 0,
                three = 0,
                four = 0,
                five = 0,
            };
            switch ((ex.Cells[smechenie1, smechenie2] as Excel.Range).Value2.ToString())
            {
                case "2":
                    buf.two++;
                    break;
                case "3":
                    buf.three++;
                    break;
                case "4":
                    buf.four++;
                    break;
                case "5":
                    buf.five++;
                    break;
                default:
                    break;
            }

            int stroka = smechenie1 + 1;
            while (ex.get_Range("C" + stroka).Value2 != null)
            {
                NexYacheika = (ex.Cells[stroka, smechenie1] as Excel.Range).Value2.ToString();
                if (PrevYacheika.Equals(NexYacheika))
                    buf.plan++;
                else
                {
                    PrevYacheika = NexYacheika;
                    SchoolMas.Add(buf);
                    buf = new School
                    {
                        name = NexYacheika,
                        plan = 1,
                        two = 0,
                        three = 0,
                        four = 0,
                        five = 0,
                    };
                }
                if ((ex.get_Range("G" + stroka).Value2 != null))
                    switch ((ex.Cells[stroka, smechenie2] as Excel.Range).Value2.ToString())
                    {
                        case "2":
                            buf.two++;
                            break;
                        case "3":
                            buf.three++;
                            break;
                        case "4":
                            buf.four++;
                            break;
                        case "5":
                            buf.five++;
                            break;
                        default:
                            break;
                    }

                stroka++;
            }
            SchoolMas.Add(buf);
            ex.Quit();
        }
        private void Outpute()
        {
            Excel.Application ex = new Excel.Application();

            Excel.Workbook Book = ex.Workbooks.Add();
            Book.SaveAs(pathOut);
            ex.Workbooks.Open(pathOut);

            ex.Cells[1, 1] = "Гбоу";
            ex.Cells[1, 2] = "План";
            ex.Cells[1, 3] = "Фактическое колво";
            ex.Cells[1, 4] = "2";
            ex.Cells[1, 5] = "3";
            ex.Cells[1, 6] = "4";
            ex.Cells[1, 7] = "5";

            for (int i = 0; i < SchoolMas.Count; i++)
            {
                ex.Cells[i + 3, 1] = SchoolMas[i].name;
                ex.Cells[i + 3, 2] = SchoolMas[i].plan;
                ex.Cells[i + 3, 3] = SchoolMas[i].two + SchoolMas[i].three + SchoolMas[i].four + SchoolMas[i].five;
                ex.Cells[i + 3, 4] = SchoolMas[i].two;
                ex.Cells[i + 3, 5] = SchoolMas[i].three;
                ex.Cells[i + 3, 6] = SchoolMas[i].four;
                ex.Cells[i + 3, 7] = SchoolMas[i].five;
            }
            ex.Quit();
        }
        private void FilePath_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fbd = new CommonOpenFileDialog();
            fbd.ShowDialog();
            fbd.Title = "Выберете файл со списком учащихся";
            if (fbd.IsCollectionChangeAllowed())
            {
                pathIn = fbd.FileName;
                pathOut = pathIn.Remove(pathIn.Length - 5, 5) + "(Статистика).xlsx";
                CalculateBut.IsEnabled = true;
            }
        }
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            if (pathIn != null)
                System.Diagnostics.Process.Start("explorer", pathIn.Substring(0, pathIn.LastIndexOf(@"\")));
            this.Close();
        }
    }
}
