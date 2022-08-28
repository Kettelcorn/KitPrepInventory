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
        private Worksheet ws;
        private Dictionary<string, string> kits;
        public MainWindow()
        {
            InitializeComponent();
            readExcel();
        }

        
        private void button_Search(object send, RoutedEventArgs e)
        {
            string boxNumber = textBox.Text + "";
            if (kits.ContainsKey(boxNumber)) {
                int number = Int32.Parse(kits[boxNumber]);
                textBlock.Text = ws.Cells[number, 3].Value + " " + ws.Cells[number, 2].Value + " in "
                    + ws.Cells[number, 5].Value + " (" + ws.Cells[number, 4].Value + ")";
            } else
            {
                textBlock.Text = "That box does not exist dumb dumb";
            }
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
            ws = wb.Worksheets[2];

            kits = new Dictionary<string, string>();
            int counter = 2; 
            while (ws.Cells[counter, 1].Value + "" != "")
            {
                kits.Add(ws.Cells[counter, 1].Value + "", counter + "");
                counter++;
            }
        }

        private void button_Search2(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
