using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
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
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Net;
 
namespace HelpAudit
{
    public partial class Form1 : Form
    {
        public Form1()
        {            
            InitializeComponent();
            int lk = 2;
            int spb = 2;
            int tk = 2;
            int cup = 2;
            int cop = 2;
            int gpgp = 2;
            int dubp = 2;
            int ok = 2;
            int suab = 2;
            int ditis = 2;
            int mop = 2;
            int oko = 2;
            int sup = 2;
            int de = 2;
            int kd = 2;
            XLWorkbook workbook = new XLWorkbook("C:\\ttt\\test.xlsx");
            var worksheet = workbook.Worksheet("Result");
    //        MessageBox.Show(worksheet.Cell(2,15 ).Value.ToString());//1386 /29

            for( int i = 2;i<=1386;i++)
            {
                string printL = "=СЦЕПИТЬ(";
                string printK = "=СЦЕПИТЬ(";
                string printM = "=СЦЕПИТЬ(";
                for (int j = 15; j < 30; j++)
                {
                    switch (j)
                    {
                        case (15):
                            {
                                if(worksheet.Cell(i, j).Value.ToString()== "v")
                                {
                                    printL += "ЛК!E" + lk+";СИМВОЛ(10);";
                                    printK += "ЛК!F" + lk + ";СИМВОЛ(10);";
                                    printM += "ЛК!G" + lk + ";СИМВОЛ(10);";
                                    lk++;
                                }
                                break;
                            }
                        case 16:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "СБП!E" + spb + ";СИМВОЛ(10);";
                                    printK += "СБП!F" + spb + ";СИМВОЛ(10);";
                                    printM += "СБП!G" + spb + ";СИМВОЛ(10);";
                                    spb++;
                                }
                                break;
                            }
                        case 17:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ТК!E" + tk + ";СИМВОЛ(10);";
                                    printK += "ТК!F" + tk + ";СИМВОЛ(10);";
                                    printM += "ТК!G" + tk + ";СИМВОЛ(10);";
                                    tk++;
                                }
                                break;
                            }
                        case 18:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ЦУП!E" + cup + ";СИМВОЛ(10);";
                                    printK += "ЦУП!F" + cup + ";СИМВОЛ(10);";
                                    printM += "ЦУП!G" + cup + ";СИМВОЛ(10);";
                                    cup++;
                                }
                                break;
                            }
                        case 19:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "КОП!E" + cop + ";СИМВОЛ(10);";
                                    printK += "КОП!F" + cop + ";СИМВОЛ(10);";
                                    printM += "КОП!G" + cop + ";СИМВОЛ(10);";
                                    cop++;
                                }
                                break;
                            }
                        case 20:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ГПГП!E" + gpgp + ";СИМВОЛ(10);";
                                    printK += "ГПГП!F" + gpgp + ";СИМВОЛ(10);";
                                    printM += "ГПГП!G" + gpgp + ";СИМВОЛ(10);";
                                    gpgp++;
                                }
                                break;
                            }
                        case 21:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ДУБП!E" + dubp + ";СИМВОЛ(10);";
                                    printK += "ДУБП!F" + dubp + ";СИМВОЛ(10);";
                                    printM += "ДУБП!G" + dubp + ";СИМВОЛ(10);";
                                    dubp++;
                                }
                                break;
                            }
                        case 22:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ОК!E" + ok + ";СИМВОЛ(10);";
                                    printK += "ОК!F" + ok + ";СИМВОЛ(10);";
                                    printM += "ОК!G" + ok + ";СИМВОЛ(10);";
                                    ok++;
                                }
                                break;
                            }
                        case 23:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "СУАБ!E" + suab + ";СИМВОЛ(10);";
                                    printK += "СУАБ!F" + suab + ";СИМВОЛ(10);";
                                    printM += "СУАБ!G" + suab + ";СИМВОЛ(10);";
                                    suab++;
                                }
                                break;
                            }
                        case 24:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ДИТИС!E" + ditis + ";СИМВОЛ(10);";
                                    printK += "ДИТИС!F" + ditis + ";СИМВОЛ(10);";
                                    printM += "ДИТИС!G" + ditis + ";СИМВОЛ(10);";
                                    ditis++;
                                }
                                break;
                            }
                        case 25:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "МОП!E" + mop + ";СИМВОЛ(10);";
                                    printK += "МОП!F" + mop + ";СИМВОЛ(10);";
                                    printM += "МОП!G" + mop + ";СИМВОЛ(10);";
                                    mop++;  
                                }
                                break;
                            }
                        case 26:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ОКО!E" + oko + ";СИМВОЛ(10);";
                                    printK += "ОКО!F" + oko + ";СИМВОЛ(10);";
                                    printM += "ОКО!G" + oko + ";СИМВОЛ(10);";
                                    oko++;  
                                }
                                break;
                            }
                        case 27:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "СУП!E" + sup + ";СИМВОЛ(10);";
                                    printK += "СУП!F" + sup + ";СИМВОЛ(10);";
                                    printM += "СУП!G" + sup + ";СИМВОЛ(10);";
                                    sup++;
                                }
                                break;
                            }
                        case 28:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "ДЭ!E" + de + ";СИМВОЛ(10);";
                                    printK += "ДЭ!F" + de + ";СИМВОЛ(10);";
                                    printM += "ДЭ!G" + de + ";СИМВОЛ(10);";
                                    de++;
                                }
                                break;
                            }
                        
                        case 29:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "КД!E" + kd + ";СИМВОЛ(10);";
                                    printK += "КД!F" + kd + ";СИМВОЛ(10);";
                                    printM += "КД!G" + kd + ";СИМВОЛ(10);";
                                    kd++;
                                }
                                break;
                            }
                    }
                    
                }
                printL += ")";
                printK += ")";
                printM += ")";
                worksheet.Cell(i, 11).Value = printL;
                worksheet.Cell(i, 12).Value = printK;
                worksheet.Cell(i, 13).Value = printM;
            }
            workbook.SaveAs("C:\\ttt\\test2.xlsx");
        }
    }
}
