using System;
using System.Collections;
using Win = System.Windows;
using Win32 = Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using Forms = System.Windows.Forms;

namespace GetPriceExcel
{
    delegate void UpdateProgressBarDelegate(Win.DependencyProperty dp, object value);
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Win.Window
    {
        private string work, filename, volume, material, smr;
        private List<GridRows> kindOfWorks = new List<GridRows>();
        //private List<GridRows> workToInsert = new List<GridRows>();
        Excel.Application app = null;
        Excel.Workbook wBook = null;
        Excel.Worksheet wSheet = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void DirectoryBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            Win32.OpenFileDialog opf = new Win32.OpenFileDialog();
            opf.Filter = "Файлы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
            opf.ShowDialog();
            filename = opf.FileName;
            try
            {
                app = new Excel.Application();
                wBook = app.Workbooks.Open(filename);
                wSheet = (Excel.Worksheet)wBook.Sheets[1];
            }
            catch (Exception ex)
            {
                Win.MessageBox.Show(ex.Message.ToString(), "An error occured in opening application file: ");
            }
            try
            {
                for (int row = 12; row < 376; row++)
                {
                    bool index = false;
                    work = wSheet.Cells[row, 3].Text;
                    if (work.Trim() != "")
                    {
                        work = work.ToLower().Trim();
                        while (work.Contains("  ")) { work = work.Replace("  ", " "); }
                        foreach (GridRows mem in kindOfWorks)
                        {
                            if (mem.Works.Equals(work)) index = true;
                        }
                        if (!index)
                        {
                            kindOfWorks.Add(new GridRows(work, volume, material, smr));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Win.MessageBox.Show(ex.Message.ToString(), "An error occured in getting data from file: ");
            }
            worksGrid.ItemsSource = kindOfWorks;
        }

        private void InsertBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            int k = worksGrid.Items.Count;
            for (int i = 0; i < k; i++)
            {
                GridRows example = ((GridRows)worksGrid.Items[i]);
                string conc = example.Works.Trim().ToLower();
                for (int row = 12; row < 376; row++)
                {
                    work = wSheet.Cells[row, 3].Text;
                    if (work.Trim() != "")
                    {
                        work = work.ToLower().Trim();
                        while (work.Contains("  ")) { work = work.Replace("  ", " "); }
                        if (conc.Equals(work)) 
                        {
                            wSheet.Cells[row, 5].Value = example.Materials;
                            wSheet.Cells[row, 6].Value = example.Smr;
                        }
                    }
                }
            }
            wBook.Save();
            Win.MessageBox.Show("Task completed", "System message: ");
        }

        private void ExitBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            wBook.Close();
            app.Quit();
            this.Close();
        }

        private void CancelBtn_Click(object sender, Win.RoutedEventArgs e)
        {

        }

    }
}
