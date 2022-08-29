using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace KitPrepProgram
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook wb;
        private Worksheet ws1;
        private Worksheet ws2;
        public MainWindow()
        {
            InitializeComponent();
            readExcel();
        }

        
        private void button_Search(object send, RoutedEventArgs e)
        {
            int counter = 2;
            string answer = "";
            while (ws2.Cells[counter, 1].Value + "" != "" && answer == "")
            {
                if (ws2.Cells[counter, 1].Value + "" == textBox.Text)
                    answer = ws2.Cells[counter, 3].Value + " " + ws2.Cells[counter, 2].Value + " in " +
                    ws2.Cells[counter, 5].Value + " (" + ws2.Cells[counter, 4].Value + ")";
                counter++;
            }
            if (answer == "") textBlock.Text = "That box does not exist dumb dumb";
            else textBlock.Text = answer;
        }

        private void button_Search2(object sender, RoutedEventArgs e)
        {
            int counter = 4;
            string answer = "";
            textBlock2.Text = "";
            while (ws1.Cells[counter, 1].Value + "" != "")
            {
                if (((string)ws1.Cells[counter, 2].Value + "").ToLower().Contains(textBox2.Text.ToLower()))
                {
                    answer = ws1.Cells[counter, 1].Value;
                    textBlock2.Text += answer + " (" + ws1.Cells[counter, 2].Value + ")\n";
                }
                counter++;  
            }
            if (answer == "") textBlock2.Text = "That item does not exist";
        }

        private void page1_Click(object send, RoutedEventArgs e)
        {
            Page1 pg = new Page1();
            this.Content = pg;
        }

            private void readExcel()
        {
            string path = @"C:\Users\Jackson Kettel\Documents\Coding\KitPrepInventory\Inventory Tracker.xlsx";
            //string path = @"C:\Users\Kette\Documents\GitHub\KitPrepInventory\Inventory Tracker.xlsx";

            excel = new Microsoft.Office.Interop.Excel.Application();
            wb = excel.Workbooks.Open(path);
            ws1 = wb.Worksheets[1];
            ws2 = wb.Worksheets[2];  
        }
    }
}
