using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Aspose.Cells;

namespace PekoDice
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void RunQ1(object sender, RoutedEventArgs e)
        {
            string SavePath = Environment.CurrentDirectory + @"/DiceNage1.xlsx";
            Workbook WB1 = new Workbook();
            Worksheet Sheet1 = WB1.Worksheets[0];
            long NTimes;
            bool IsNumber = long.TryParse(Q1Number.Text, out NTimes);
            if (!IsNumber)
            {
                return;
            }
            string[,] TmpData = new string[NTimes + 1, NTimes + 2];
            for (int i = 1; i <= NTimes; i++)
            {
                TmpData[0, i] = "Dice" + i.ToString();
            }
            TmpData[0, NTimes + 1] = "Average";
            for (long i = 1; i <= NTimes; i++)
            {
                decimal Average = 0;
                TmpData[i, 0] = i.ToString() + " Times";
                for (long j = 1; j <= i; j++)
                {
                    Random Rng = new Random(Guid.NewGuid().GetHashCode());
                    int Num = Rng.Next(1, 7);
                    TmpData[i, j] = Num.ToString();
                    Average += Num;
                }
                Average /= i;
                TmpData[i, NTimes + 1] = Average.ToString();
            }
            Sheet1.Cells.ImportArray(TmpData, 0, 0);
            Sheet1.AutoFitColumns();
            Sheet1.AutoFitRows();
            try
            {
                WB1.Save(SavePath);
                Console.WriteLine("Create file successful.");
            }
            catch (Exception)
            {
                Console.WriteLine("File has opened. Please close the file and retry.");
            }
        }

        private void RunQ2(object sender, RoutedEventArgs e)
        {
            string SavePath = Environment.CurrentDirectory + @"/DiceNage2.xlsx";
            Workbook WB1 = new Workbook();
            Worksheet Sheet1 = WB1.Worksheets[0];
            int NDice;
            bool IsNumber = int.TryParse(Q2Number.Text, out NDice);
            if (!IsNumber)
            {
                return;
            }
            string[,] TmpTitle = new string[3, 100 + 1];
            TmpTitle[0, 0] = NDice.ToString() + " Dices";
            TmpTitle[1, 0] = "Sum";
            TmpTitle[2, 0] = "Average";
            double[,] TmpData = new double[2, 100];

            for (int i = 1; i <= 100; i++)
            {
                TmpTitle[0, i] = i.ToString() + " Times";
            }
            NKakeru100(NDice, ref TmpData, 0);
            Sheet1.Cells.ImportArray(TmpTitle, 0, 0);
            Sheet1.Cells.ImportArray(TmpData, 1, 1);
            Sheet1.AutoFitColumns();
            Sheet1.AutoFitRows();
            try
            {
                WB1.Save(SavePath);
                Console.WriteLine("Create file successful.");
            }
            catch (Exception)
            {
                Console.WriteLine("File has opened. Please close the file and retry.");
            }
        }

        private void NKakeru100(int N, ref double[,] TmpData, int Rowi)
        {

            for (int i = 0; i < 100; i++)
            {
                double Average = 0;
                for (int j = 0; j < N; j++)
                {
                    Random Rng = new Random(Guid.NewGuid().GetHashCode());
                    Average += Rng.Next(1, 7);
                }
                TmpData[Rowi, i] = Average;
                Average /= N;
                TmpData[Rowi + 1, i] = Average;
            }
        }
    }
}
