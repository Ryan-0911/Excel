using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Excel
{
    public partial class Form1 : Form
    {
        private void btn讀取單一儲存格_Click(object sender, EventArgs e)
        {
            ExcelManager mgr = new ExcelManager();
            mgr.OpenFile(@"C:\Users\S2239002\Desktop\test1.xlsx", 1);
            MessageBox.Show(mgr.ReadCell(1, 3));
        }

        private void btn寫入單一儲存格_Click(object sender, EventArgs e)
        {
            ExcelManager mgr = new ExcelManager();
            // 原始檔案
            mgr.OpenFile(@"C:\Users\S2239002\Desktop\test1.xlsx", 1);
            mgr.WriteToCell(1, 3, "test");
            // 新檔案
            mgr.CreateNewFile();
            mgr.WriteToCell(1, 3, "test");
            mgr.SaveAs(@"C:\Users\S2239002\Desktop\0120.xlsx");
        }

        private void btn讀取多重儲存格_Click(object sender, EventArgs e)
        {
            ExcelManager mgr = new ExcelManager();
            mgr.OpenFile(@"C:\Users\S2239002\Desktop\0120.xlsx", 1);
            string[,] result = mgr.ReadRange(1, 1, 5, 3);
        }

        private void btn寫入多重儲存格_Click(object sender, EventArgs e)
        {
            ExcelManager mgr = new ExcelManager();
            mgr.OpenFile(@"C:\Users\S2239002\Desktop\0120.xlsx", 1);
            string[,] result = mgr.ReadRange(1, 1, 5, 3);
            mgr.CreateNewFile();
            mgr.WriteRange(1, 1, 3, 3, result);
            mgr.SaveAs(@"C:\Users\S2239002\Desktop\0212.xlsx");
        }

        public Form1()
        {
            InitializeComponent();
        }
    }
}
