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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Search(object send, RoutedEventArgs e)
        {
            readExcel();
        }

            private void readExcel()
        {
            string path = @"C:\Users\Jackson Kettel\Documents\Coding\ARLN Kit Prep Inventory.xlsx";

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
            int counter = 2; 
            while (ws.Cells[counter, 1].Value + "" != textBox.Text)
            {
                counter++;
            }
            if (ws.Cells[counter, 1].Value + "" == "")
                textBlock.Text = "That box does not exist";
            else
                textBlock.Text = 
                  ws.Cells[counter, 3].Value + " "
                + ws.Cells[counter, 2].Value + " in "
                + ws.Cells[counter, 5].Value + " ("
                + ws.Cells[counter, 4].Value + ")";
        }
    }
}
