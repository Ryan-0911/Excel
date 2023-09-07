using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal class ExcelManager
    {
        Application excelApp = new Application();
        Workbook excelWB;
        Worksheet excelWS;
        Range range;

        // 開啟工作簿
        public void OpenFile(string pathAndFileName, int sheet)
        {
            excelWB = excelApp.Workbooks.Open(pathAndFileName);
            excelWS = excelWB.Worksheets[sheet];
        }

        // 建立工作簿
        public void CreateNewFile()
        {
            excelWB = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            excelWS = excelWB.Worksheets[1];
        }

        // 新增工作表
        public void CreateNewSheet()
        {
            excelWB.Worksheets.Add(After: excelWS);
        }

        // 選擇工作表
        public void SelectWorksheet(int num)
        {
            this.excelWS = excelWB.Worksheets[num];
        }

        // 刪除工作頁
        public void DeleteWorksheet(int num)
        {
            excelWB.Worksheets[num].Delete();
        }

        // 從單一儲存格讀取資料
        public string ReadCell(int i, int j)
        {
            if (excelWS.Cells[i, j].Value2 != null)
            {
                return excelWS.Cells[i, j].Value2;
            }
            else
            {
                return "can't read any data";
            }
        }

        // 將資料寫入單一儲存格
        public void WriteToCell(int i, int j, string s)
        {
            excelWS.Cells[i, j].Value2 = s;
        }

        // 從多重儲存格讀取資料
        public string[,] ReadRange(int startx, int starty, int endx, int endy)
        {
            Range range = (Range)excelWS.Range[excelWS.Cells[startx, starty], excelWS.Cells[endx, endy]];
            object[,] holder = range.Value2;
            string[,] stringHolder = new string[endx - startx, endy - starty];
            for (int i = 1; i <= endx - startx; i++)
            {
                for (int j = 1; j <= endy - starty; j++)
                {
                    stringHolder[i - 1, j - 1] = holder[i, j].ToString();
                }
            }
            return stringHolder;
        }

        // 將資料寫入多重儲存格
        public void WriteRange(int startx, int starty, int endx, int endy, string[,] writeString)
        {
            Range range = (Range)excelWS.Range[excelWS.Cells[startx, starty], excelWS.Cells[endx, endy]];
            range.Value2 = writeString;
        }

        // 保護工作表 (改為非保護模式無需密碼)
        public void ProtectSheet()
        {
            excelWS.Protect();
        }

        // 保護工作表 (改為非保護模式需要密碼)
        public void ProtectSheet(string password)
        {
            excelWS.Protect(password);
        }

        // 非保護工作表 (改為保護模式無需密碼)
        public void UnProtectSheet()
        {
            excelWS.Unprotect();
        }

        // 非保護工作表 (改為保護模式需要密碼)
        public void UnProtectSheet(string password)
        {
            excelWS.Unprotect(password);
        }

        // 存檔
        public void Save()
        {
            excelWB.Save();
        }

        // 另存新檔
        public void SaveAs(string s)
        {
            excelWB.SaveAs(s);
        }

        // 釋放資源
        public void Close()
        {
            excelWB.Close();
            excelApp.Quit();
        }
    }
}
