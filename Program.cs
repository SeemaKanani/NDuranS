﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel; 

namespace ConsoleApplication40
{
    class Program
    {
        static public Dictionary<string, SecDataRow> GSECDict;
        static public Dictionary<string, string> GSECNameDict;

        static Excel.Application excelApp;
        static Excel.Workbook excelWorkbook;
        static Excel.Sheets excelSheets;
        static Excel.Worksheet excelWorksheet;

        static int NumberFields = 7;
        static int PollSleepTime = 500;

        static void Main(string[] args)
        {
            //System.AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);

            string tickerDirPath = @"C:\Program Files (x86)\GILTS\CCIG02P1\XXXXXXCF\";
            string fileName = DateTime.Now.ToString("yyyyMMdd") + "_aplnrcvr.log";
            //string tickerFilePath = @"C:\Users\kananip\Downloads\Seema\20161122_aplnrcvrhrsh\test.txt";
            string tickerFilePath = Path.Combine(tickerDirPath, fileName);
            string excelFilePath = @"c:\n\test.xlsx";
            string excelWorksheetName = "GSEC";
            
            GSECDict = new Dictionary<string, SecDataRow>();
            GSECNameDict = new Dictionary<string, string>();


            InitExcel(excelFilePath,excelWorksheetName);
            StartMonitorFile(tickerFilePath);

            Console.WriteLine("<FAIL> Most likely '{0}' does not exist !",tickerFilePath);
            Console.WriteLine("Press a key to EXIT");
            Console.Read();
        }

        static void OnProcessExit(object sender, EventArgs e)
        {
            Console.WriteLine("I'm out of here");
            //excelWorksheet.
            excelWorkbook.Close(0);
            excelApp.Quit();

            //ReleaseObje
        }

        static void WriteToExcel()
        {
            object[,] arr = new object[GSECDict.Count, SecDataRow.RowLen];
            int counter = 0 ;
            foreach(SecDataRow kv in GSECDict.Values ) 
            {
                if (kv.isZero())
                    continue;
                string[] temp = kv.GetRow();
                for (int i = 0; i < SecDataRow.RowLen; i++)
                    arr[counter, i] = temp[i];
                counter ++; 
            }
            Console.WriteLine("<INFO> Updating Excel. {0} Row(s) affected",counter);
            string endCellIndex = /*((char)('A' + SecDataRow.RowLen-1)).ToString() "AQ"*/ "AB"+counter;
            Excel.Range excelCell = (Excel.Range)excelWorksheet.get_Range("A1", endCellIndex);
            excelCell.Value2 = arr;

        }

        static void InitExcel(string excelFilePath,string excelWorksheetName)
        {
            Console.WriteLine("<INFO> opening excel file : " + excelFilePath + ":" + excelWorksheetName);
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            string workbookPath = excelFilePath; // @"c:\n\test.xlsx";
            excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            excelSheets = excelWorkbook.Worksheets;
            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(excelWorksheetName);
            Console.WriteLine("<SUCCESS> excel file successfully opened");
            Console.WriteLine();
        }

        static void StartMonitorFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using (var sr = new StreamReader(fs))
                {

                    // We will be using polling .. 
                    while (true)
                    {
                        string currentBlock = sr.ReadToEnd();
                        if (!String.IsNullOrEmpty(currentBlock))
                        {
                            string[] currentLines = currentBlock.Split('\n');
                            if (currentLines.Length > 0)
                            {
                                if (Parse(currentLines))
                                    WriteToExcel();
                            }
                        }
                        Thread.Sleep(PollSleepTime);
                    }
                }
            }
        }

        static bool Parse(string[] lines)
        {
            /*
             * 1000 :    20161122163835.529375000
               ~BUYX~$~0~$~0~$~0
               ~SELL~
                4994~1~105.59~6.4062~100000000~1~0~
                4994~2~105.595~6.4054~100000000~1~0~
                4994~3~105.5975~6.405~100000000~1~0~
             * */
            bool result = false;
            foreach (string line in lines)
            {
                if (line.Contains("~SELL~") && line.Contains("~BUYX~") && line.Contains("PUSH"))
                {
                    int indexBuyx = line.IndexOf("~BUYX~");
                    int indexSell = line.IndexOf("~SELL~");

                    if (indexSell > indexBuyx)
                    {
                        string stringBuyx = line.Substring(indexBuyx + 6, indexSell - indexBuyx - 6);
                        string stringSell = line.Substring(indexSell + 6);

                        string[] arrBuyx = stringBuyx.Split('$');
                        string[] arrSell = stringSell.Split('$');

                        string data = arrBuyx[0];
                        {
                            string[] arrData = data.Split('~');
                            int arrDataLen = arrData.Length;
                            for (int i = 0; i < arrDataLen; i += NumberFields)
                            {
                                if (i + NumberFields < arrDataLen)
                                {
                                    string[] secData = new string[NumberFields];
                                    Array.Copy(arrData, i, secData, 0, NumberFields);
                                    SecData temp = new SecData(secData);

                                    if (GSECNameDict.ContainsKey(temp.id))
                                        temp.name = GSECNameDict[temp.id];

                                    string key = temp.id;

                                    if (temp.number == "1")
                                    {
                                        if (GSECDict.ContainsKey(key))
                                        {
                                            GSECDict[key].Buyx = temp;
                                        }
                                        else
                                        {
                                            SecDataRow sdr = new SecDataRow();
                                            sdr.Buyx = temp;
                                            GSECDict.Add(key, sdr);
                                        }
                                        result = true;
                                    }
                                }
                           }
                       }

                        foreach (string datum in arrBuyx)
                        {
                            if (datum.Contains("~TAKN~"))
                            {
                                string[] splitArr = datum.Split('~');
                                if (splitArr.Length >= 19)
                                {
                                    int index = -1;
                                    for (int i = 15; i < splitArr.Length; i++)
                                    {
                                        if (splitArr[i] == "TAKN")
                                        {
                                            index = i;
                                            break;
                                        }
                                    }

                                    if (index == -1)
                                        continue;

                                    string key = splitArr[index - 14];
                                    string[] inputArr = new string[16];
                                    Array.Copy(splitArr, index-13, inputArr, 0, 16);
                                    if (GSECDict.ContainsKey(key))
                                    {
                                        GSECDict[key].Taken = inputArr;
                                        GSECDict[key].isGiven = false;
                                    }
                                    else
                                    {
                                        SecDataRow sdr = new SecDataRow();
                                        sdr.Taken = inputArr;
                                        sdr.isGiven = false;
                                        GSECDict.Add(key, sdr);
                                    }
                                }
                            }

                            if (datum.Contains("~GIVN~"))
                            {
                                string[] splitArr = datum.Split('~');
                                if (splitArr.Length >= 19)
                                {
                                    int index = -1;
                                    for (int i = 15; i < splitArr.Length; i++)
                                    {
                                        if (splitArr[i] == "GIVN")
                                        {
                                            index = i;
                                            break;
                                        }
                                    }

                                    if (index == -1)
                                        continue;
                                    string key = splitArr[index-14];
                                    string[] inputArr = new string[16];
                                    Array.Copy(splitArr, index-13, inputArr, 0, 16);
                                    if (GSECDict.ContainsKey(key))
                                    {
                                        GSECDict[key].Given = inputArr;
                                        GSECDict[key].isGiven = true;
                                    }
                                    else
                                    {
                                        SecDataRow sdr = new SecDataRow();
                                        sdr.Given = inputArr;
                                        sdr.isGiven = true;
                                        GSECDict.Add(key, sdr);
                                    }
                                }

                            }
                        }

                        foreach (string datum in arrSell)
                        {
                            if (datum.Contains("~TAKN~"))
                            {
                                string[] splitArr = datum.Split('~');
                                if (splitArr.Length >= 19)
                                {
                                    int index = -1;
                                    for (int i = 15; i < splitArr.Length; i++)
                                    {
                                        if (splitArr[i] == "TAKN")
                                        {
                                            index = i;
                                            break;
                                        }
                                    }

                                    if (index == -1)
                                        continue;
                                    string key = splitArr[index-14];
                                    string[] inputArr = new string[16];
                                    Array.Copy(splitArr, index - 13, inputArr, 0, 16);
                                    if (GSECDict.ContainsKey(key))
                                    {
                                        GSECDict[key].Taken = inputArr;
                                        GSECDict[key].isGiven = false;
                                    }
                                    else
                                    {
                                        SecDataRow sdr = new SecDataRow();
                                        sdr.Taken = inputArr;
                                        sdr.isGiven = false;
                                        GSECDict.Add(key, sdr);
                                    }
                                }
                            }

                            if (datum.Contains("~GIVN~"))
                            {
                                string[] splitArr = datum.Split('~');
                                if (splitArr.Length >= 19)
                                {
                                    int index = -1;
                                    for (int i = 15; i < splitArr.Length; i++)
                                    {
                                        if (splitArr[i] == "GIVN")
                                        {
                                            index = i;
                                            break;
                                        }
                                    }

                                    if (index == -1)
                                        continue;

                                    string key = splitArr[index-14];
                                    string[] inputArr = new string[16];
                                    Array.Copy(splitArr, index-13, inputArr, 0, 16);
                                    if (GSECDict.ContainsKey(key))
                                    {
                                        GSECDict[key].Given = inputArr;
                                        GSECDict[key].isGiven = true;
                                    }
                                    else
                                    {
                                        SecDataRow sdr = new SecDataRow();
                                        sdr.Given = inputArr;
                                        sdr.isGiven = true;
                                        GSECDict.Add(key, sdr);
                                    }
                                }

                            }
                        }

                          data = arrSell[0];
                          {
                            string[] arrData = data.Split('~');
                            int arrDataLen = arrData.Length;
                            for (int i = 0; i < arrDataLen; i += NumberFields)
                            {
                                if (i + NumberFields < arrDataLen)
                                {
                                    string[] secData = new string[NumberFields];
                                    Array.Copy(arrData, i, secData, 0, NumberFields);
                                    SecData temp = new SecData(secData);

                                    if (GSECNameDict.ContainsKey(temp.id))
                                        temp.name = GSECNameDict[temp.id];
                                    string key = temp.id;

                                    if (GSECDict.ContainsKey(key))
                                    {
                                        GSECDict[key].Sell = temp;
                                    }
                                    else
                                    {
                                        SecDataRow sdr = new SecDataRow();
                                        sdr.Sell = temp;
                                        GSECDict.Add(key, sdr);
                                    }
                                    result = true;
                                }
                            }
                        }
                    }
                }
                // we want to find share names.. 
                    //1701 :    Y~6557~3686~T1XX~T1XX~COUP~IN0020109016~IN0020109016~08.01 POSTAL LIFE INS GOI SPL SEC 2021~1589~0~8.01~20160930~20170331~20170930~20200930~9~53~127~180~180~180~FIXD~0~~~0~1~6559~3687~T1XX~T1XX~COUP~IN0020109024~IN0020109024~08.08 POSTAL LIFE INS GOI SPL SEC 2023~2319~0~8.08~20160930~20170331~20170930~20220930~13~53~127~180~180~180~FIXD~0~~~0~1~4876~2982~T1XX~T1XX~COUP~IN0020089077~IN0020089077~08.00 OIL MKTG COS GOI SB 2026~3407~0~8~20160923~20170323~20170923~20250923~19~60~120~180~180~180~FIXD~0~~~0~1~4994~3028~T1XX~T1XX~COUP~IN0020090034~IN0020090034~07.35 GOVT. STOCK 2024~2768~901680220000~7.35~20160622~20161222~20170622~20231222~16~151~29~180~180~180~FIXD~0~~~0~1~11520~5820~T1XX~T1XX~COUP~IN0020160050~IN0020160050~06.84 GOVT. STOCK 2022~2217~130000000000~6.84~20160912~20161219~20170619~20220619~13~71~26~97~180~180~FIXD~0~~~0~1~2462~2447~T1XX~T1XX~COUP~IN0020060037~IN0020060037~08.20 GOVT. STOCK 2022~1910~576323300000~8.2~20160815~20170215~20170815~20210815~11~98~82~180~180~180~FIXD~0~~~0~1~7943~4280~T1XX~T1XX~COUP~IN0020130012~IN0020130012~07.16 GOVT. STOCK 2023~2369~771000000000~7.16~20161120~20170520~20171120~20221120~13~3~177~180~180~180~FIXD~0~~~0~1~5695~3326~T1XX~T1XX~COUP~IN0020100031~IN0020100031~08.30 GOVT. STOCK 2040~8622~900000000000~8.3~20160702~20170102~20170702~20400102~48~141~39~180~180~180~FIXD~0~~~0~1~5042~3049~T1XX~T1XX~COUP~IN0020090042~IN0020090042~06.90 GOVT. STOCK 2019~962~450000000000~6.9~20160713~20170113~20170713~20190113~6~130~50~180~180~180~FIXD~0~~~0~1~8902~4684~T1XX~T1XX~COUP~IN0020140011~IN0020140011~08.60 GOVT. STOCK 2028~4209~840000000000~8.6~20160602~20161202~20170602~20271202~24~171~9~180~180~180~FIXD~0~~~0~1~8289~4422~T1XX~T1XX~COUP~IN0020130053~IN0020130053~09.20 GOVT. STOCK 2030~5059~618845500000~9.2~20160930~20170330~20170930~20300330~28~53~127~180~180~180~FIXD~0~~~0~1~2195~2280~T1XX~T1XX~COUP~IN0020020171~IN0020020171~06.35 GOVT. STOCK 2020~1135~610000000000~6.35~20160702~20170102~20170702~20190702~7~141~39~180~180~180~FIXD~0~~~0~1~2194~2279~T1XX~T1XX~COUP~IN0020010065~IN0020010065~10.03 GOVT. STOCK 2019~989~60000000000~10.03~20160809~20170209~20170809~20190209~6~104~76~180~180~180~FIXD~0~~~0~1~2193~2278~T1XX~T1XX~COUP~IN0020030048~IN0020030048~06.05 GOVT. STOCK 2019~931~110000000000~6.05~20160612~20161212~20170612~20181212~6~161~19~180~180~180~FIXD~0~~~0~1~2192~2277~T1XX~T1XX~COUP~IN0020030097~IN0020030097~05.64 GOVT. STOCK 2019~770~100000000000~5.64~20160702~20170102~20170702~20180702~5~141~39~180~180~180~FIXD~0~~~0~1~2191~2276~T1XX~T1XX~COUP~IN0019980286~IN0019980286~12.60 GOVT. STOCK 2018~730~126318800000~12.6~20161123~20170523~20171123~20180523~4~0~180~180~180~180~FIXD~0~~~0~1~2190~2275~T1XX~T1XX~COUP~IN0020030063~IN0020030063~05.69 GOVT. STOCK 2018~671~161300000000~5.69~20160925~20170325~20170925~20180325~4~58~122~180~180~180~FIXD~0~~~0~1~2189~2274~T1XX~T1XX~COUP~IN0020010024~IN0020010024~10.45 GOVT. STOCK 2018~523~37160000000~10.45~20161030~20170430~20171030~20171030~3~23~157~180~180~180~FIXD~0~~~0~1~$~N~  :    1 :    PUSH
                else if (line.Contains("T1XX~T1XX~COUP"))
                {
                    string[] parts = line.Split('~');
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (parts[i] == "COUP")
                        {
                            string ID = parts[i - 4];
                            string Name = parts[i + 3];
                            if (!GSECNameDict.ContainsKey(ID))
                            {
                                GSECNameDict.Add(ID, Name);
                            }
                        }
                    }
                }
            }

            return result;
        }

        static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            Console.WriteLine("<FATAL> Caught exception !");
            Console.WriteLine(e.ExceptionObject.ToString());
            Console.WriteLine("Press a key to EXIT");
            Console.Read();
            Environment.Exit(1);
        }
    }

    class SecData
    {
        public string id, number, price, yield, amount, x, y,name;
        public SecData()
        {
        }

        public SecData(string[] arrData)
        {
           id =     arrData[0];
           number = arrData[1];
           price =  arrData[2];
           yield =  arrData[3];
           amount = arrData[4];
           x =      arrData[5];
           y =      arrData[6];
        }

        public string[] GetArray()
        {
            string[] result = new string[7];

            result[0] = id;
            result[1] = name;
            result[2] = price;
            result[3] = yield;
            result[4] = amount;
            result[5] = x;
            result[6] = y;

            return result;
        }

        public bool isZero()
        {
            bool result = true;
            string[] data = GetArray();
            for (int i = 2; i < data.Length; i++)
            {
                if (data[i] == null)
                    return true;
                if (!(data[i].Trim() == "0" ))
                { result = false; break; }
            }
            return result;
        }
    }

    class SecDataRow
    {
        public SecData Sell, Buyx;
        public string[] Taken, Given;
        public bool isGiven;
        //public string id;
        public static int RowLen = 12 + 16;
        public SecDataRow()
        {
            Sell = new SecData();
            Buyx = new SecData();
            Taken = new string[16];
            Given = new string[16];
            isGiven = false;
        }
        public string[] GetRow()
        {
            int numberSecData = 2;
            int numberCommonFields = 2;
            int rowLen = (numberSecData * 7 - (numberSecData - 1) * numberCommonFields)+Taken.Length;
            string[] row = new string[rowLen];

            string[] SellData = Sell.GetArray();
            string[] BuyxData = Buyx.GetArray();
            int index = 0;
            Array.Copy(SellData, 0, row, 0, index+= SellData.Length);

            for (int i = 2; i < BuyxData.Length; i++, index++)
            {
                row[index] = BuyxData[i];
            }
            string[] testArr = new string[16];

            testArr = isGiven ? Given : Taken;

            for (int i = 0; i < testArr.Length; i++, index++)
            {
                row[index] = testArr[i];
            }

            //for (int i = 0; i < Taken.Length; i++, index++)
            //{
            //    row[index] = Taken[i];
            //}


            return row;
        }

        public bool isZero()
        {
              //  return Sell.isZero() && Buyx.isZero();
            return false;
        }
    }
}
