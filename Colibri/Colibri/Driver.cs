using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Colibri
{
    class Driver
    {
        public bool run(string orderLoc,string transLoc,string resultLoc)
        {
                Microsoft.Office.Interop.Excel.Application transXL=null;
                Microsoft.Office.Interop.Excel._Workbook transWB=null;
                Microsoft.Office.Interop.Excel._Worksheet transSheet=null;

                Microsoft.Office.Interop.Excel.Application orderXl=null;
                Microsoft.Office.Interop.Excel._Workbook orderBook=null;
                Microsoft.Office.Interop.Excel._Worksheet orderSheet=null;

                Microsoft.Office.Interop.Excel.Application resultXl=null;
                Microsoft.Office.Interop.Excel._Workbook resultBook=null;
                Microsoft.Office.Interop.Excel._Worksheet resultSheet=null;
            int count=1;
            while (File.Exists(resultLoc+"\\Result"+count+".xlsx"))
                count++;

            try
            {
                transXL = new Microsoft.Office.Interop.Excel.Application();
                transWB = transXL.Workbooks.Open(transLoc);
                transSheet = (Microsoft.Office.Interop.Excel._Worksheet)transWB.ActiveSheet;
              
               
                orderXl = new Microsoft.Office.Interop.Excel.Application();
                orderBook = orderXl.Workbooks.Open(orderLoc);
                orderSheet = (Microsoft.Office.Interop.Excel._Worksheet)orderBook.ActiveSheet;
               
                resultXl = new Microsoft.Office.Interop.Excel.Application();
                //resultBook = resultXl.Workbooks.Open(resultLoc + "\\Result" + count + ".xlsx");
                resultBook = resultXl.Workbooks.Add();
                resultSheet = (Microsoft.Office.Interop.Excel._Worksheet)resultBook.ActiveSheet;
               
                Dictionary<string, School> schools;
                Dictionary<string, int[]> itemTotal;

                ReadFiles(transSheet, orderSheet, out schools, out itemTotal);
                SchoolInfo(schools, resultSheet);
                itemInfo(itemTotal, resultSheet);
                
                resultSheet.SaveAs(resultLoc + "\\Result" + count + ".xlsx");

                transWB.Close();
                transXL.Quit();
                orderBook.Close();
                orderXl.Quit();
                resultBook.Close();
                resultXl.Quit();
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error with Processing");
                transWB.Close();
                transXL.Quit();
                orderBook.Close();
                orderXl.Quit();
                resultBook.Close();
                resultXl.Quit();
                return false;
            }
  }

        private static void ReadFiles(_Worksheet transSheet, _Worksheet orderSheet, out Dictionary<string, School> schools, out Dictionary<string, int[]> itemTotal)
        {
            var lastTransRow = transSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            var transOBJ = (object[,])transSheet.UsedRange.Value2;

            var lastOrderRow = orderSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            var orderOBJ = (object[,])orderSheet.UsedRange.Value2;

            int orderNamePos = 18;
            int orderQuantPos = 17;
            int orderNumberPos = 1;
            int orderSKUPos = 21;
            int orderSchoolPos = 15;

            string schoolName = "UNKNOWNSCHOOL";
            string orderName;
            string orderItemSize;
            int orderSKU;
            int orderNum;
            int orderQuant;

            double amount;
            double fee;
            double net;

            schools = new Dictionary<string, School>();
            itemTotal = new Dictionary<string, int[]>();

            for (int i = 2; i < lastOrderRow + 1; i++)
            {
                try
                {
                    orderNum = Convert.ToInt32(Clean(orderOBJ[i, orderNumberPos]));
                }
                catch
                {
                    orderNum = -1;
                    System.Windows.Forms.MessageBox.Show("Order: Invalid Order Number Found On Line " + i);
                }
                if (!string.IsNullOrEmpty(Convert.ToString(orderOBJ[i, orderSchoolPos])))
                {
                    schoolName = Convert.ToString(orderOBJ[i, orderSchoolPos]);
                }
                orderName = Dice(Convert.ToString(orderOBJ[i, orderNamePos]));
                try
                {
                    orderSKU = Convert.ToInt32(Clean(orderOBJ[i, orderSKUPos]));
                }
                catch
                {
                    orderSKU = -1;
                    System.Windows.Forms.MessageBox.Show("Order: Invalid SKU Found On Line " + i);
                }               
                try
                {
                    orderQuant = Convert.ToInt32(Clean(orderOBJ[i, orderQuantPos]));
                }
                catch
                {
                    orderQuant = -1;
                    System.Windows.Forms.MessageBox.Show("Order: Invalid Order Quantity Found On Line "+i);
                }
                orderItemSize = FindSize(Convert.ToString(orderOBJ[i, orderNamePos]));

               if (!schools.ContainsKey(schoolName))
                    schools.Add(schoolName, new School(schoolName));

                if (orderName.Contains("Organic"))
                {
                    schools[schoolName].AddItem(Convert.ToString(orderOBJ[i, orderNamePos]) , orderSKU, orderQuant, orderItemSize);
                    schools[schoolName].AddOrderNum(orderNum);
                    AddQuant(itemTotal, Convert.ToString(orderOBJ[i, orderNamePos]), string.Empty, orderSKU, orderQuant);
                }
                else if (orderItemSize.Equals("Set"))
                {
                    int last3 = orderSKU-8000;
                    schools[schoolName].AddItem(orderName + " " + "Small", last3+1000, orderQuant,  "Small");
                    schools[schoolName].AddItem(orderName + " " + "Large", last3+2000, orderQuant,  "Large");
                    schools[schoolName].AddItem(orderName + " " + "Medium", last3+9000, orderQuant,  "Medium");
                    schools[schoolName].AddOrderNum(orderNum);

                    AddQuant(itemTotal, orderName, "Small", last3 + 1000, orderQuant);
                    AddQuant(itemTotal, orderName, "Large", last3 + 2000, orderQuant);
                    AddQuant(itemTotal, orderName, "Medium", last3 + 9000, orderQuant);
                }
                else
                {
                    schools[schoolName].AddItem(orderName + " " + orderItemSize, orderSKU, orderQuant, orderItemSize);
                    schools[schoolName].AddOrderNum(orderNum);
                    AddQuant(itemTotal, orderName, orderItemSize, orderSKU, orderQuant);
                }
            }
            orderNumberPos = 3;
            int amountPos = 9;
            int feePos = 10;
            int netPos = 11;
            for (int i = 2; i < lastTransRow+1; i++)
            {
                try
                {
                    orderNum = Convert.ToInt32(Clean(transOBJ[i, orderNumberPos]));
                    amount = Clean(transOBJ[i, amountPos]) != "" ? Convert.ToDouble(Clean(transOBJ[i, amountPos])) : 0;
                    fee = Clean(transOBJ[i, feePos]) != "" ? Convert.ToDouble(Clean(transOBJ[i, feePos])) : 0;
                    net = Clean(transOBJ[i, netPos]) != "" ? Convert.ToDouble(Clean(transOBJ[i, netPos])) : 0;
                    foreach (string currSch in schools.Keys)
                    {
                        if (schools[currSch].HasOrder(orderNum))
                            schools[currSch].AddMoney(orderNum, net, fee, amount);
                    }
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Transaction: Invalid Number Found On Line "+i);
                }
            }
        }
        private static void AddQuant(Dictionary<string,int[]> itemTotal,string orderName,string orderItemSize, int sku, int orderQuant)
        {
            if (!itemTotal.ContainsKey(orderName + " " + orderItemSize))
            {
                itemTotal.Add(orderName + " " + orderItemSize, new int[] { 0,  sku });
            }
            itemTotal[orderName + " " + orderItemSize][0] = itemTotal[orderName + " " + orderItemSize][0] + orderQuant;
        }
        private static string Dice(string fullName)
        {
            string chopName = fullName;
            chopName =chopName.Replace(" Large","");
            chopName = chopName.Replace(" Regular", "");
            chopName = chopName.Replace(" Medium", "");
            chopName = chopName.Replace(" Set", "");
            chopName = chopName.Replace(" Mini", "");
            chopName = chopName.Replace(" Medium", "");
            chopName = chopName.Replace(" Snack", "");
            chopName = chopName.Replace(" Reusable", "");
            chopName = chopName.Replace(" Small", "");
            chopName = chopName.Replace(" -", "");
            chopName = chopName.Replace(" Bag", "");
            return chopName;
        }

        private static string FindSize(string orderName)
        {
            string result = "Unknown Size";
            orderName = orderName.ToLower();

            if (orderName.Contains("large"))
                result = "Large";
            else if (orderName.Contains("small"))
                result = "Small";            
            else if (orderName.Contains("set"))
                result = "Set";
            else if (orderName.Contains("regular"))
                result = "Regular";
            else if (orderName.Contains("organic"))
                result = "Organic";
            else if (orderName.Contains("straw"))
                result = "Straw";
            else if (orderName.Contains("mini"))
                result = "Mini";
            else if (orderName.Contains("medium"))
                result = "Medium";
            else
            {

            }

            return result;
        }

        private static string Clean(object obj)
        {
            string result = "";
            string input = Convert.ToString(obj);
            foreach (char letter in input)
            {
                if (((letter >= '0') && (letter <= '9'))||(letter=='.'))
                    result += letter;
            }
            return result;
        }

        private void itemInfo(Dictionary<string, int[]> itemTotal, _Worksheet resultSheet)
        {
            int itemTotalStart = 20;
            int itemWrite = 5;

            resultSheet.Cells[1, itemTotalStart] = "Total Small";
            resultSheet.Cells[1, itemTotalStart + 1] = "Total Regular";
            resultSheet.Cells[1, itemTotalStart + 2] = "Total Large";
            resultSheet.Cells[1, itemTotalStart + 3] = "Total Medium";
            resultSheet.Cells[1, itemTotalStart + 4] = "Total Mini";
            resultSheet.Cells[1, itemTotalStart + 5] = "Total Wash Clothes";
            resultSheet.Cells[1, itemTotalStart + 6] = "Total Straws";

            resultSheet.Cells[2, itemTotalStart] = GetItemTotal(itemTotal,"Small");
            resultSheet.Cells[2, itemTotalStart + 1] = GetItemTotal(itemTotal,"Regular");
            resultSheet.Cells[2, itemTotalStart + 2] = GetItemTotal(itemTotal, "Large");
            resultSheet.Cells[2, itemTotalStart + 3] = GetItemTotal(itemTotal, "Medium");
            resultSheet.Cells[2, itemTotalStart + 4] = GetItemTotal(itemTotal, "Mini");
            resultSheet.Cells[2, itemTotalStart + 5] = GetItemTotal(itemTotal, "Organic");
            resultSheet.Cells[2, itemTotalStart + 6] = GetItemTotal(itemTotal, "Straw");

            resultSheet.Cells[itemWrite, itemTotalStart] = "Item Name";
            resultSheet.Cells[itemWrite, itemTotalStart + 1] = "Item Number";
            resultSheet.Cells[itemWrite++, itemTotalStart + 2] = "Item SKU";
            while (itemTotal.Count!=0)
            {
                string key="None";
                int kingNumber = 0;
                foreach (string name in itemTotal.Keys)
                {
                    if (itemTotal[name][0]>kingNumber)
                    {
                        kingNumber = itemTotal[name][0];
                        key = name;
                    }
                }
                resultSheet.Cells[itemWrite, itemTotalStart] = key;
                resultSheet.Cells[itemWrite, itemTotalStart+1] = kingNumber;
                resultSheet.Cells[itemWrite, itemTotalStart + 2] = itemTotal[key][1];

                itemTotal.Remove(key);
                itemWrite++;
            }
            resultSheet.UsedRange.Columns.AutoFit();
            resultSheet.UsedRange.HorizontalAlignment= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }
        private int GetItemTotal(Dictionary<string, int[]> itemTotal, string target)
        {
            int total = 0;
            foreach (string name in itemTotal.Keys)
            {
                if (name.Contains(target))
                    total += itemTotal[name][0];
            }
            return total;
        }
        private static void SchoolInfo(Dictionary<string, School> schools, _Worksheet resultSheet)
        {
            double[] mon;

            int totalWrite = 2;
            int moneyWrite = 2;
            int itemWrite = 2;

            int itemTot;

            int orderTotalsStart = 1;
            int totalsColStart = 8;
            int itemSchoolStart = 14;

            string itemType="Small";

            XlRgbColor[] Colours = new XlRgbColor[] {XlRgbColor.rgbLightGreen, XlRgbColor.rgbLightYellow, XlRgbColor.rgbLavender, XlRgbColor .rgbLightBlue, XlRgbColor .rgbLightSalmon};
            int schoolColourIndex=0;

            resultSheet.Cells[1, orderTotalsStart] = "School Name";
            resultSheet.Cells[1, orderTotalsStart + 1] = "Order Number";
            resultSheet.Cells[1, orderTotalsStart + 2] = "Net";
            resultSheet.Cells[1, orderTotalsStart + 3] = "Fee";
            resultSheet.Cells[1, orderTotalsStart + 4] = "Amount";

            resultSheet.Cells[1, totalsColStart] = "School Name";
            resultSheet.Cells[1, totalsColStart + 1] = "Total Net";
            resultSheet.Cells[1, totalsColStart + 2] = "Total Fee";
            resultSheet.Cells[1, totalsColStart + 3] = "Total Amount";
            resultSheet.Cells[1, totalsColStart + 4] = "Number of Orders";

            resultSheet.Cells[1, itemSchoolStart] = "School Name";
            resultSheet.Cells[1, itemSchoolStart + 1] = "Item Name";
            resultSheet.Cells[1, itemSchoolStart + 2] = "Number of Items";
            resultSheet.Cells[1, itemSchoolStart + 3] = "SKU";
            resultSheet.Cells[1, itemSchoolStart + 4] = "Items per type";


            foreach (string schoolName in schools.Keys)
            {
                mon = schools[schoolName].GetAllTotals();
                resultSheet.Cells[totalWrite, totalsColStart] = schoolName;
                resultSheet.Cells[totalWrite, totalsColStart+1] = mon[0];
                resultSheet.Cells[totalWrite, totalsColStart+2] = mon[1];
                resultSheet.Cells[totalWrite, totalsColStart + 3] = mon[2];
                resultSheet.Cells[totalWrite, totalsColStart + 4] = mon[3];
                resultSheet.Range[resultSheet.Cells[totalWrite, totalsColStart], resultSheet.Cells[totalWrite, totalsColStart + 4]].Interior.Color=Colours[schoolColourIndex];
                totalWrite++;

                resultSheet.Cells[moneyWrite, orderTotalsStart] = schoolName;
                Dictionary<int, double[]> money = schools[schoolName].GetMoney();
                foreach (int orderNum in money.Keys)
                {
                    mon = money[orderNum];
                    resultSheet.Cells[moneyWrite, orderTotalsStart+1] = orderNum;
                    resultSheet.Cells[moneyWrite, orderTotalsStart + 2] = mon[0];
                    resultSheet.Cells[moneyWrite, orderTotalsStart + 3] = mon[1];
                    resultSheet.Cells[moneyWrite, orderTotalsStart + 4] = mon[2];
                    resultSheet.Range[resultSheet.Cells[moneyWrite, orderTotalsStart], resultSheet.Cells[moneyWrite, orderTotalsStart + 4]].Interior.Color = Colours[schoolColourIndex];

                    moneyWrite++;
                }

                resultSheet.Cells[itemWrite, itemSchoolStart] = schoolName;
                itemType = "Small";
                itemTot = 0;
                foreach (string[] item in schools[schoolName].GetSortItems())
                {                    
                    resultSheet.Cells[itemWrite, itemSchoolStart + 1] = item[0];
                    resultSheet.Cells[itemWrite, itemSchoolStart + 2] = item[1];
                    resultSheet.Cells[itemWrite, itemSchoolStart + 3] = item[2];
                    resultSheet.Range[resultSheet.Cells[itemWrite, itemSchoolStart], resultSheet.Cells[itemWrite, itemSchoolStart+4]].Interior.Color = Colours[schoolColourIndex];


                    if (itemType.Equals(item[0].Substring(item[0].LastIndexOf(" ")+1)))
                        itemTot += Convert.ToInt32(item[1]);
                    else
                    {
                        resultSheet.Cells[itemWrite - 1, itemSchoolStart + 4] = itemTot.ToString();
                        itemType = item[0].Substring(item[0].LastIndexOf(" ")+1);
                        itemTot = Convert.ToInt32(item[1]);
                    }

                    itemWrite++;                   
                }
                resultSheet.Cells[itemWrite - 1, itemSchoolStart + 4] = itemTot.ToString();

                schoolColourIndex = (schoolColourIndex+1)%Colours.Length;
            }
        }
    }
}
