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
        private Dictionary<int, int> kits;
        public MainWindow()
        {
            InitializeComponent();
            readExcel();
        }

        
        private void button_Search(object send, RoutedEventArgs e)
        {
            int boxNumber = Int32.Parse(textBox.Text);
            if (kits.ContainsKey(boxNumber)) {
                textBlock.Text = ws.Cells[kits[boxNumber], 3] + ws.Cells[kits[boxNumber], 2] + " in "
                    + ws.Cells[kits[boxNumber], 5] + " (" + ws.Cells[kits[boxNumber], 4] + ")";
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
            string path = @"C:\Users\Jackson Kettel\Documents\Coding\ARLN Kit Prep Inventory.xlsx";
            //string path = @"C:\Users\Kette\Documents\GitHub\KitPrepInventory\Inventory Tracker.xlsx";

            excel = new Microsoft.Office.Interop.Excel.Application();
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];

            kits = new Dictionary<int, int>();
            int counter = 2; 
            while (ws.Cells[counter, 1].Value + "" != "")
            {
                kits.Add(ws.Cells[counter, 1].Value, counter);
                counter++;
            }
        }

        private void button_Search2(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
