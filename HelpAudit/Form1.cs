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

            for( int i = 2;i<1386;i++)
            {
                string printL = "=СЦЕПИТЬ(";
                string printK = "=СЦЕПИТЬ(";
                string printM = "=СЦЕПИТЬ(";
                string printN = "=СЦЕПИТЬ(";
                string printAD = "=СЦЕПИТЬ(";
                string printAE = "=СЦЕПИТЬ(";
                string printAF = "=СЦЕПИТЬ(";
                string printAG = "=СЦЕПИТЬ(";
                string printAH = "=СЦЕПИТЬ(";
                for (int j = 15; j < 30; j++)
                {
                    switch (j)
                    {
                        case (15):
                            {
                                if(worksheet.Cell(i, j).Value.ToString()== "v")
                                {
                                    printL += "\"ЛК:\";"+"ЛК!E" + lk+";СИМВОЛ(10);";
                                    printK += "\"ЛК:\";" + "ЛК!F" + lk + ";СИМВОЛ(10);";
                                    printM += "\"ЛК:\";" + "ЛК!G" + lk + ";СИМВОЛ(10);";
                                    printN += "\"ЛК:\";" + "ЛК!H" + lk + ";СИМВОЛ(10);";
                                    printAD += "\"ЛК:\";" + "ЛК!I" + lk + ";СИМВОЛ(10);";
                                    printAE += "\"ЛК:\";" + "ЛК!J" + lk + ";СИМВОЛ(10);";
                                    printAF += "\"ЛК:\";" + "ЛК!K" + lk + ";СИМВОЛ(10);";
                                    printAG += "\"ЛК:\";" + "ЛК!L" + lk + ";СИМВОЛ(10);";
                                    printAH += "\"ЛК:\";" + "ЛК!M" + lk + ";СИМВОЛ(10);";
                                    lk++;
                                }
                                break;
                            }
                        case 16:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"СБП: \";" + "СБП!E" + spb + ";СИМВОЛ(10);";
                                    printK += "\"СБП: \";" + "СБП!F" + spb + ";СИМВОЛ(10);";
                                    printM += "\"СБП: \";" + "СБП!G" + spb + ";СИМВОЛ(10);";
                                    printN += "\"СБП: \";" + "СБП!H" + spb + ";СИМВОЛ(10);";
                                    printAD += "\"СБП: \";" + "СБП!I" + spb + ";СИМВОЛ(10);";
                                    printAE += "\"СБП: \";" + "СБП!J" + spb + ";СИМВОЛ(10);";
                                    printAF += "\"СБП:\";" + "СБП!K" + spb + ";СИМВОЛ(10);";
                                    printAG += "\"СБП: \";" + "СБП!L" + spb + ";СИМВОЛ(10);";
                                    printAH += "\"СБП: \";" + "СБП!M" + spb + ";СИМВОЛ(10);";
                                    spb++;
                                }
                                break;
                            }
                        case 17:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ТК: \";" + "ТК!E" + tk + ";СИМВОЛ(10);";
                                    printK += "\"ТК: \";" + "ТК!F" + tk + ";СИМВОЛ(10);";
                                    printM += "\"ТК: \";" + "ТК!G" + tk + ";СИМВОЛ(10);";
                                    printN += "\"ТК: \";" + "ТК!H" + tk + ";СИМВОЛ(10);";
                                    printAD += "\"ТК: \";" + "ТК!I" + tk + ";СИМВОЛ(10);";
                                    printAE += "\"ТК: \";" + "ТК!J" + tk + ";СИМВОЛ(10);";
                                    printAF += "\"ТК: \";" + "ТК!K" + tk + ";СИМВОЛ(10);";
                                    printAG += "\"ТК: \";" + "ТК!L" + tk + ";СИМВОЛ(10);";
                                    printAH += "\"ТК: \";" + "ТК!M" + tk + ";СИМВОЛ(10);";
                                    tk++;
                                }
                                break;
                            }
                        case 18:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ЦУП: \";" + "ЦУП!E" + cup + ";СИМВОЛ(10);";
                                    printK += "\"ЦУП: \";" + "ЦУП!F" + cup + ";СИМВОЛ(10);";
                                    printM += "\"ЦУП: \";" + "ЦУП!G" + cup + ";СИМВОЛ(10);";
                                    printN += "\"ЦУП: \";" + "ЦУП!H" + cup + ";СИМВОЛ(10);";
                                    printAD += "\"ЦУП: \";" + "ЦУП!I" + cup + ";СИМВОЛ(10);";
                                    printAE += "\"ЦУП: \";" + "ЦУП!J" + cup + ";СИМВОЛ(10);";
                                    printAF += "\"ЦУП: \";" + "ЦУП!K" + cup + ";СИМВОЛ(10);";
                                    printAG += "\"ЦУП: \";" + "ЦУП!L" + cup + ";СИМВОЛ(10);";
                                    printAH += "\"ЦУП: \";" + "ЦУП!M" + cup + ";СИМВОЛ(10);";
                                    cup++;
                                }
                                break;
                            }
                        case 19:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"КОП: \";" + "КОП!E" + cop + ";СИМВОЛ(10);";
                                    printK += "\"КОП: \";" + "КОП!F" + cop + ";СИМВОЛ(10);";
                                    printM += "\"КОП: \";" + "КОП!G" + cop + ";СИМВОЛ(10);";
                                    printN += "\"КОП: \";" + "КОП!H" + cop + ";СИМВОЛ(10);";
                                    printAD += "\"КОП: \";" + "КОП!I" + cop + ";СИМВОЛ(10);";
                                    printAE += "\"КОП: \";" + "КОП!J" + cop + ";СИМВОЛ(10);";
                                    printAF += "\"КОП: \";" + "КОП!K" + cop + ";СИМВОЛ(10);";
                                    printAG += "\"КОП: \";" + "КОП!L" + cop + ";СИМВОЛ(10);";
                                    printAH += "\"КОП: \";" + "КОП!M" + cop + ";СИМВОЛ(10);";
                                    cop++;
                                }
                                break;
                            }
                        case 20:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ГПГП: \";" + "ГПГП!E" + gpgp + ";СИМВОЛ(10);";
                                    printK += "\"ГПГП: \";" + "ГПГП!F" + gpgp + ";СИМВОЛ(10);";
                                    printM += "\"ГПГП: \";" + "ГПГП!G" + gpgp + ";СИМВОЛ(10);";
                                    printN += "\"ГПГП: \";" + "ГПГП!H" + gpgp + ";СИМВОЛ(10);";
                                    printAD += "\"ГПГП: \";" + "ГПГП!I" + gpgp + ";СИМВОЛ(10);";
                                    printAE += "\"ГПГП: \";" + "ГПГП!J" + gpgp + ";СИМВОЛ(10);";
                                    printAF += "\"ГПГП: \";" + "ГПГП!K" + gpgp + ";СИМВОЛ(10);";
                                    printAG += "\"ГПГП: \";" + "ГПГП!L" + gpgp + ";СИМВОЛ(10);";
                                    printAH += "\"ГПГП: \";" + "ГПГП!M" + gpgp + ";СИМВОЛ(10);";
                                    gpgp++;
                                }
                                break;
                            }
                        case 21:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"КУБП: \";" + "КУБП!E" + dubp + ";СИМВОЛ(10);";
                                    printK += "\"КУБП: \";" + "КУБП!F" + dubp + ";СИМВОЛ(10);";
                                    printM += "\"КУБП: \";" + "КУБП!G" + dubp + ";СИМВОЛ(10);";
                                    printN += "\"КУБП: \";" + "КУБП!H" + dubp + ";СИМВОЛ(10);";
                                    printAD += "\"КУБП: \";" + "КУБП!I" + dubp + ";СИМВОЛ(10);";
                                    printAE += "\"КУБП: \";" + "КУБП!J" + dubp + ";СИМВОЛ(10);";
                                    printAF += "\"КУБП: \";" + "КУБП!K" + dubp + ";СИМВОЛ(10);";
                                    printAG += "\"КУБП: \";" + "КУБП!L" + dubp + ";СИМВОЛ(10);";
                                    printAH += "\"КУБП: \";" + "КУБП!M" + dubp + ";СИМВОЛ(10);";
                                    dubp++;
                                }
                                break;
                            }
                        case 22:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ОК: \";" + "ОК!E" + ok + ";СИМВОЛ(10);";
                                    printK += "\"ОК: \";" + "ОК!F" + ok + ";СИМВОЛ(10);";
                                    printM += "\"ОК: \";" + "ОК!G" + ok + ";СИМВОЛ(10);";
                                    printN += "\"ОК: \";" + "ОК!H" + ok + ";СИМВОЛ(10);";
                                    printAD += "\"ОК: \";" + "ОК!I" + ok + ";СИМВОЛ(10);";
                                    printAE += "\"ОК: \";" + "ОК!J" + ok + ";СИМВОЛ(10);";
                                    printAF += "\"ОК: \";" + "ОК!K" + ok + ";СИМВОЛ(10);";
                                    printAG += "\"ОК: \";" + "ОК!L" + ok + ";СИМВОЛ(10);";
                                    printAH += "\"ОК: \";" + "ОК!M" + ok + ";СИМВОЛ(10);";
                                    ok++;
                                }
                                break;
                            }
                        case 23:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"СУАБ: \";" + "СУАБ!E" + suab + ";СИМВОЛ(10);";
                                    printK += "\"СУАБ: \";" + "СУАБ!F" + suab + ";СИМВОЛ(10);";
                                    printM += "\"СУАБ: \";" + "СУАБ!G" + suab + ";СИМВОЛ(10);";
                                    printN += "\"СУАБ: \";" + "СУАБ!H" + suab + ";СИМВОЛ(10);";
                                    printAD += "\"СУАБ: \";" + "СУАБ!I" + suab + ";СИМВОЛ(10);";
                                    printAE += "\"СУАБ: \";" + "СУАБ!J" + suab + ";СИМВОЛ(10);";
                                    printAF += "\"СУАБ: \";" + "СУАБ!K" + suab + ";СИМВОЛ(10);";
                                    printAG += "\"СУАБ: \";" + "СУАБ!L" + suab + ";СИМВОЛ(10);";
                                    printAH += "\"СУАБ: \";" + "СУАБ!M" + suab + ";СИМВОЛ(10);";
                                    suab++;
                                }
                                break;
                            }
                        case 24:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ДИТИС: \";" + "ДИТИС!E" + ditis + ";СИМВОЛ(10);";
                                    printK += "\"ДИТИС: \";" + "ДИТИС!F" + ditis + ";СИМВОЛ(10);";
                                    printM += "\"ДИТИС: \";" + "ДИТИС!G" + ditis + ";СИМВОЛ(10);";
                                    printN += "\"ДИТИС: \";" + "ДИТИС!H" + ditis + ";СИМВОЛ(10);";
                                    printAD += "\"ДИТИС: \";" + "ДИТИС!I" + ditis + ";СИМВОЛ(10);";
                                    printAE += "\"ДИТИС: \";" + "ДИТИС!J" + ditis + ";СИМВОЛ(10);";
                                    printAF += "\"ДИТИС: \";" + "ДИТИС!K" + ditis + ";СИМВОЛ(10);";
                                    printAG += "\"ДИТИС: \";" + "ДИТИС!L" + ditis + ";СИМВОЛ(10);";
                                    printAH += "\"ДИТИС: \";" + "ДИТИС!M" + ditis + ";СИМВОЛ(10);";
                                    ditis++;
                                }
                                break;
                            }
                        case 25:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"МОП: \";" + "МОП!E" + mop + ";СИМВОЛ(10);";
                                    printK += "\"МОП: \";" + "МОП!F" + mop + ";СИМВОЛ(10);";
                                    printM += "\"МОП: \";" + "МОП!G" + mop + ";СИМВОЛ(10);";
                                    printN += "\"МОП: \";" + "МОП!H" + mop + ";СИМВОЛ(10);";
                                    printAD += "\"МОП: \";" + "МОП!I" + mop + ";СИМВОЛ(10);";
                                    printAE += "\"МОП: \";" + "МОП!J" + mop + ";СИМВОЛ(10);";
                                    printAF += "\"МОП: \";" + "МОП!K" + mop + ";СИМВОЛ(10);";
                                    printAG += "\"МОП: \";" + "МОП!L" + mop + ";СИМВОЛ(10);";
                                    printAH += "\"МОП: \";" + "МОП!M" + mop + ";СИМВОЛ(10);";
                                    mop++;  
                                }
                                break;
                            }
                        case 26:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ОКО: \";" + "ОКО!E" + oko + ";СИМВОЛ(10);";
                                    printK += "\"ОКО: \";" + "ОКО!F" + oko + ";СИМВОЛ(10);";
                                    printM += "\"ОКО: \";" + "ОКО!G" + oko + ";СИМВОЛ(10);";
                                    printN += "\"ОКО: \";" + "ОКО!H" + oko + ";СИМВОЛ(10);";
                                    printAD += "\"ОКО: \";" + "ОКО!I" + oko + ";СИМВОЛ(10);";
                                    printAE += "\"ОКО: \";" + "ОКО!J" + oko + ";СИМВОЛ(10);";
                                    printAF += "\"ОКО: \";" + "ОКО!K" + oko + ";СИМВОЛ(10);";
                                    printAG += "\"ОКО: \";" + "ОКО!L" + oko + ";СИМВОЛ(10);";
                                    printAH += "\"ОКО: \";" + "ОКО!M" + oko + ";СИМВОЛ(10);";
                                    oko++;  
                                }
                                break;
                            }
                        case 27:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"СУП: \";" + "СУП!E" + sup + ";СИМВОЛ(10);";
                                    printK += "\"СУП: \";" + "СУП!F" + sup + ";СИМВОЛ(10);";
                                    printM += "\"СУП: \";" + "СУП!G" + sup + ";СИМВОЛ(10);";
                                    printN += "\"СУП: \";" + "СУП!H" + sup + ";СИМВОЛ(10);";
                                    printAD += "\"СУП: \";" + "СУП!I" + sup + ";СИМВОЛ(10);";
                                    printAE += "\"СУП: \";" + "СУП!J" + sup + ";СИМВОЛ(10);";
                                    printAF += "\"СУП: \";" + "СУП!K" + sup + ";СИМВОЛ(10);";
                                    printAG += "\"СУП: \";" + "СУП!L" + sup + ";СИМВОЛ(10);";
                                    printAH += "\"СУП: \";" + "СУП!M" + sup + ";СИМВОЛ(10);";
                                    sup++;
                                }
                                break;
                            }
                        case 28:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"ДЭ: \";" + "ДЭ!E" + de + ";СИМВОЛ(10);";
                                    printK += "\"ДЭ: \";" + "ДЭ!F" + de + ";СИМВОЛ(10);";
                                    printM += "\"ДЭ: \";" + "ДЭ!G" + de + ";СИМВОЛ(10);";
                                    printN += "\"ДЭ: \";" + "ДЭ!H" + de + ";СИМВОЛ(10);";
                                    printAD += "\"ДЭ: \";" + "ДЭ!I" + de + ";СИМВОЛ(10);";
                                    printAE += "\"ДЭ: \";" + "ДЭ!J" + de + ";СИМВОЛ(10);";
                                    printAF += "\"ДЭ: \";" + "ДЭ!K" + de + ";СИМВОЛ(10);";
                                    printAG += "\"ДЭ: \";" + "ДЭ!L" + de + ";СИМВОЛ(10);";
                                    printAH += "\"ДЭ: \";" + "ДЭ!M" + de + ";СИМВОЛ(10);";
                                    de++;
                                }
                                break;
                            }
                        
                        case 29:
                            {
                                if (worksheet.Cell(i, j).Value.ToString() == "v")
                                {
                                    printL += "\"КД: \";" + "КД!E" + kd + ";СИМВОЛ(10);";
                                    printK += "\"КД: \";" + "КД!F" + kd + ";СИМВОЛ(10);";
                                    printM += "\"КД: \";" + "КД!G" + kd + ";СИМВОЛ(10);";
                                    printN += "\"КД: \";" + "КД!H" + kd + ";СИМВОЛ(10);";
                                    printAD += "\"КД: \";" + "КД!I" + kd + ";СИМВОЛ(10);";
                                    printAE += "\"КД: \";" + "КД!J" + kd + ";СИМВОЛ(10);";
                                    printAF += "\"КД: \";" + "КД!K" + kd + ";СИМВОЛ(10);";
                                    printAG += "\"КД: \";" + "КД!L" + kd + ";СИМВОЛ(10);";
                                    printAH += "\"КД: \";" + "КД!M" + kd + ";СИМВОЛ(10);";
                                    kd++;
                                }
                                break;
                            }
                    }
                    
                }
                printL += ")";
                printK += ")";
                printM += ")";
                printN += ")";
                printAD += ")";
                printAE += ")";
                printAF += ")";
                printAG += ")";
                printAH += ")";
                worksheet.Cell(i, 11).Value = printL;
                worksheet.Cell(i, 12).Value = printK;
                worksheet.Cell(i, 13).Value = printM;
                worksheet.Cell(i, 14).Value = printN;
                worksheet.Cell(i, 30).Value = printAD;
                worksheet.Cell(i, 31).Value = printAE;
                worksheet.Cell(i, 32).Value = printAF;
                worksheet.Cell(i, 33).Value = printAG;
                worksheet.Cell(i, 34).Value = printAH;
            }
            workbook.SaveAs("C:\\ttt\\test2.xlsx"); 
        }
    }
}
