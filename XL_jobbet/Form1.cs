using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using IronXL;
using IronXL.Styles;
using System.Diagnostics;

/* 
Iterate through ranges and get values.
++++  https://automationtesting.in/get-data-from-excel-using-column-name-in-c/ ++++++
https://stackoverflow.com/questions/657131/how-to-read-data-of-an-excel-file-using-c
https://www.codeproject.com/Questions/1242209/Csharp-find-if-column-contains-a-N-if-so-then-dele

*/


namespace XL_jobbet
{
    public partial class Form1 : Form
    {
        string Firstopening = "0"; //skikt nummer
        string procent1 = "50";
        string SecondOpening = "0";
        string procent2 = "60";
        string ThirdOpening = "0";
        string procent3 = "60";
        int counter = 0;
        string wkBook = @"C:\Excel\BP2.xlsm";

        public Form1()
        {
            InitializeComponent();
            IronXL.License.LicenseKey = "IRONXL-BOARD4ALL.BIZ-121256-ED1A78B-32A539DF9-D4FF3B-NEx-R1B";

            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            killExcel();


        }

        private void button1_Click(object sender, EventArgs e)
        {


            WorkBook workbook = WorkBook.Load(wkBook);
            WorkSheet oldSheet = workbook.WorkSheets.First();
            WorkSheet sheet = workbook.GetWorkSheet("Blad1");
            var range = sheet["D14:D50"];


            //***********************************

            foreach (var cell in range)
            {
                double CCell = sheet[$"C{cell.RowIndex }"].DoubleValue;
                string DCell = sheet[$"D{cell.RowIndex }"].ToString();
                double ECell = sheet[$"E{cell.RowIndex }"].DoubleValue;
                ////MessageBox.Show(" skikt nr: " + ECell.ToString());

                //********************Iterate down the columns and read a specific cell************************//


                if (DCell.ToString().Contains("G") && CCell >= 1.99999)
                {

                    ////MessageBox.Show(" skikt nr: " + ECell.ToString() + " Innehåller: " + DCell + " Opt. Thickness = " + CCell.ToString());

                    double openingLayer = ECell + 2; // Ändra i denna om öppningen ska ske i ett anna skikt. flytta tvåan till en global variabel
                    double openingTextLayer; // Cell nummer och operation för "fyll på material och korrigera" Textgen 

                    string procentageOpening_C = "50%";
                    string procentageOpening_D = "60%";
                    string procentageOpening_E = "60%";
                    //  string SheetMachineNameM19 = "M19 (ZnS)"; // "M27 (SiO)";
                    string SheetMachineNameM20 = "M20 (ZnS)"; // "M27 (SiO)";
                    string SheetMachineNameM21 = "M21 (ZnS)"; // "M27 (SiO)";

                    // //MessageBox.Show("opening_Layer is : " + openingLayer.ToString());

                    /*
                                        if (counter == 2)
                                        {
                                            openingTextLayer = openingLayer + 7;
                                            ////MessageBox.Show("Counter value is : " + counter.ToString());
                                            ///  createOpening(openingLayer.ToString(), procentageOpening_C, "E", SheetMachineNameM19, openingTextLayer.ToString());
                                            createOpening(openingLayer.ToString(), procentageOpening_C, "E", SheetMachineNameM20, openingTextLayer.ToString());
                                            createOpening(openingLayer.ToString(), procentageOpening_C, "E", SheetMachineNameM21, openingTextLayer.ToString());                        counter++;
                                        }
                                        if (counter == 1)
                                        {
                                            openingTextLayer = openingLayer + 5;
                                            ////MessageBox.Show("Counter value is : " + counter.ToString());
                                            ///  createOpening(openingLayer.ToString(), procentageOpening_C, "D", SheetMachineNameM19, openingTextLayer.ToString());
                                            createOpening(openingLayer.ToString(), procentageOpening_C, "D", SheetMachineNameM20, openingTextLayer.ToString());
                                            createOpening(openingLayer.ToString(), procentageOpening_C, "D", SheetMachineNameM21, openingTextLayer.ToString());
                                            counter++;
                                        }
                                        if (counter == 0)
                                        {
                                            openingTextLayer = openingLayer + 3;
                                            ////MessageBox.Show("Counter value is : " + counter.ToString());
                                            ///   createOpening(openingLayer.ToString(), procentageOpening_C, "C", SheetMachineNameM19, openingTextLayer.ToString());
                                       //     createOpening(openingLayer.ToString(), procentageOpening_C, "C", SheetMachineNameM20, openingTextLayer.ToString());
                                            createOpening(openingLayer.ToString(), procentageOpening_C, "C", SheetMachineNameM21, openingTextLayer.ToString());
                                            counter++;
                                        }


                                        */

                    IEnumerable<int> collection = new List<int>() { 9, 10 }; //,  "M18 (ZnS)", "M19 (ZnS)", "M20 (ZnS)" 

                    /* */

                    // string str1 = collection.ElementAt(i);



                    if (counter == 2)
                    {
                        openingTextLayer = openingLayer + 7;
                        ////MessageBox.Show("Counter value is : " + counter.ToString());
                        ///  createOpening(openingLayer.ToString(), procentageOpening_C, "E", SheetMachineNameM19, openingTextLayer.ToString());
                        //createOpening(openingLayer.ToString(), procentageOpening_C, "E", SheetMachineNameM20, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_E, "E", collection.First(), openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_E, "E", collection.First() + 1, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_E, "E", collection.First() + 2, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_E, "E", collection.First() + 3, openingTextLayer.ToString());
                        counter++;
                    }
                    if (counter == 1)
                    {
                        openingTextLayer = openingLayer + 5;
                        ////MessageBox.Show("Counter value is : " + counter.ToString());
                        ///  createOpening(openingLayer.ToString(), procentageOpening_C, "D", SheetMachineNameM19, openingTextLayer.ToString());
                        //createOpening(openingLayer.ToString(), procentageOpening_C, "D", SheetMachineNameM20, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_D, "D", collection.First(), openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_D, "D", collection.First() + 1, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_D, "D", collection.First() + 2, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_D, "D", collection.First() + 3, openingTextLayer.ToString());

                        counter++;
                    }
                    if (counter == 0)
                    {
                        openingTextLayer = openingLayer + 3;
                        ////MessageBox.Show("Counter value is : " + counter.ToString());
                        ///   createOpening(openingLayer.ToString(), procentageOpening_C, "C", SheetMachineNameM19, openingTextLayer.ToString());
                        //     createOpening(openingLayer.ToString(), procentageOpening_C, "C", SheetMachineNameM20, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_C, "C", collection.First(), openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_C, "C", collection.First() + 1, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_C, "C", collection.First() + 2, openingTextLayer.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_C, "C", collection.First() + 3, openingTextLayer.ToString());

                        counter++;
                    }


                   // for (int i = 0; i < collection.Count(); i++)
                    //{ }


                    /*
                    if (counter == 2)
                    {
                        openingTextLayer = openingLayer + 7;
                        //MessageBox.Show("Counter value is : " + counter.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_E, "E", 10, openingTextLayer.ToString());  // Move procentage opening to a global variable.
                        counter++;
                    }
                    if (counter == 1)
                    {
                        openingTextLayer = openingLayer + 5;
                        //MessageBox.Show("Counter value is : " + counter.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_D, "D",10, openingTextLayer.ToString());
                        counter++;
                    }
                    if (counter == 0)
                    {
                        openingTextLayer = openingLayer + 3;
                        //MessageBox.Show("Counter value is : " + counter.ToString());
                        createOpening(openingLayer.ToString(), procentageOpening_C, "C", 10, openingTextLayer.ToString());
                        counter++;
                    }*/

                }




            }
            killExcel();

        }


        void createOpening(string openingLayer, string procent, string column, int SheetMachineName, string openingTextLayer)
        {

            // xlWorkbook.SaveAs2(@"C:\Excel\BP2_Edit.xlsm");



            //Create Openeing///
            if (column.Equals("C"))
            {
                //MessageBox.Show($"openong for {column}");
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Excel\BP2.xlsm", CorruptLoad: true);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Blad1"];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                xlApp.DisplayAlerts = false;






                xlWorksheet.Cells[5, "C"] = openingLayer; // öppnings skiktet

                xlWorksheet.Cells[6, "C"] = procent; // procent i öppning
                Excel._Worksheet xlSheetMachineName = xlWorkbook.Sheets[SheetMachineName];
                Excel.Range xlRange2 = xlWorksheet.UsedRange;


                xlSheetMachineName.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").ClearContents();
                xlSheetMachineName.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").MergeCells = true;



                xlSheetMachineName.get_Range($"B{openingTextLayer}").Value2 = "Fyll på material";
                xlSheetMachineName.get_Range($"B{openingTextLayer}").Font.Bold = true;
                //xlWorksheetM19.get_Range("B11", "K11").Font.Color = Color.Blue;
                xlSheetMachineName.get_Range($"B{openingTextLayer}").Cells.Interior.Color = Color.Yellow;


                xlWorkbook.SaveAs2(@"C:\Excel\BP2_Edit.xlsm");
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();

            }
            else if (column.Equals("D"))
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Excel\BP2_Edit.xlsm", CorruptLoad: true);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Blad1"];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                xlApp.DisplayAlerts = false;

                xlWorksheet.Cells[5, "D"] = openingLayer; // öppnings skiktet

                xlWorksheet.Cells[6, "D"] = procent; // procent i öppning
                Excel._Worksheet xlWorksheetM19 = xlWorkbook.Sheets[SheetMachineName];
                Excel.Range xlRange2 = xlWorksheet.UsedRange;


                xlWorksheetM19.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").ClearContents();
                xlWorksheetM19.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").MergeCells = true;



                xlWorksheetM19.get_Range($"B{openingTextLayer}").Value2 = "Fyll på material";
                xlWorksheetM19.get_Range($"B{openingTextLayer}").Font.Bold = true;
                //xlWorksheetM19.get_Range("B11", "K11").Font.Color = Color.Blue;
                xlWorksheetM19.get_Range($"B{openingTextLayer}").Cells.Interior.Color = Color.Yellow;


                xlWorkbook.SaveAs2(@"C:\Excel\BP2_Edit.xlsm");
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();

            }
            else if (column.Equals("E"))
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Excel\BP2_Edit.xlsm", CorruptLoad: true);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Blad1"];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                xlApp.DisplayAlerts = false;


                xlWorksheet.Cells[5, "E"] = openingLayer; // öppnings skiktet

                xlWorksheet.Cells[6, "E"] = procent; // procent i öppning

                Excel._Worksheet xlWorksheetM19 = xlWorkbook.Sheets[SheetMachineName];
                Excel.Range xlRange2 = xlWorksheet.UsedRange;


                xlWorksheetM19.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").ClearContents();
                xlWorksheetM19.get_Range($"B{openingTextLayer}", $"K{openingTextLayer}").MergeCells = true;



                xlWorksheetM19.get_Range($"B{openingTextLayer}").Value2 = "Fyll på material";
                xlWorksheetM19.get_Range($"B{openingTextLayer}").Font.Bold = true;
                //xlWorksheetM19.get_Range("B11", "K11").Font.Color = Color.Blue;
                xlWorksheetM19.get_Range($"B{openingTextLayer}").Cells.Interior.Color = Color.Yellow;


                xlWorkbook.SaveAs2(@"C:\Excel\BP2_Edit.xlsm");
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();
            }




        }


        void pasteOpeningText(string MachineSheetName)  // 
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Excel\BP2.xlsm", CorruptLoad: true);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[MachineSheetName];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //   Cell sourceCell = xlWorksheet.Cells["A1"];
            var range = xlWorksheet.Cells[1, 4];
            //xlRange.Formula = "= PI()";
            xlRange.NumberFormat = "0.0000";
            //xlRange.Style = style;
            xlRange.Font.Color = Color.Blue;
            xlRange.Font.Bold = true;






            /////END///////
            //close and release
            xlWorkbook.SaveAs2(@"C:\Excel\BP2_Edited_1.xlsm");
            xlWorkbook.Close();


            //quit and release
            xlApp.Quit();

            /*  */

            ///COPY Code
            /* 
             Excel.Range source = xlWorksheet.Range["A9:L9"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
             Excel.Range destination = xlWorksheet.Range["F10"];
             source.Copy(destination);
             */
        }

        private void killExcel()
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    clsProcess.Kill();
                    break;
                }
            }
        }
        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }
        private void button8_Click_1(object sender, EventArgs e)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Excel\BP2.xlsm", CorruptLoad: true);


            foreach (Excel._Worksheet worksheet in xlWorkbook.Worksheets)
            {

                MessageBox.Show(worksheet.Name, worksheet.Index.ToString());
            }

            xlWorkbook.Close();


            //quit and release
            xlApp.Quit();
        }














        private void readExcel()
        {
            /* string filePath = @"C:\Excel\BP.xlsm";
              Excel = new Excel.Application();
              Excel.Workbook wb;
              Excel.Worksheet ws;
              wb = Excel.Workbooks.Open(filePath);
              ws = wb.Worksheets[1];

              Excel.Range cell = ws.Cells["A1:A2"];
              foreach (string Result in cell.Value)
              {
                  //MessageBox.Show(Result);
              }*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Declare these two variables globally so you can access them from both
            //Button1 and Button2.
            Excel.Application objApp;
            Excel._Workbook objBook;

            Excel.Workbooks objBooks;
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range range;

            try
            {
                // Instantiate Excel and start a new workbook.
                objApp = new Excel.Application();
                objBooks = objApp.Workbooks;
                objBook = objBooks.Add(Missing.Value);
                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);

                //Get the range where the starting cell has the address
                //m_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
                range = objSheet.get_Range("A1", Missing.Value);
                range = range.get_Resize(5, 5);

                if (this.FillWithStrings.Checked == false)
                {
                    //Create an array.
                    double[,] saRet = new double[5, 5];

                    //Fill the array.
                    for (long iRow = 0; iRow < 5; iRow++)
                    {
                        for (long iCol = 0; iCol < 5; iCol++)
                        {
                            //Put a counter in the cell.
                            saRet[iRow, iCol] = iRow * iCol;
                        }
                    }

                    //Set the range value to the array.
                    range.set_Value(Missing.Value, saRet);
                }

                else
                {
                    //Create an array.
                    string[,] saRet = new string[5, 5];

                    //Fill the array.
                    for (long iRow = 0; iRow < 5; iRow++)
                    {
                        for (long iCol = 0; iCol < 5; iCol++)
                        {
                            //Put the row and column address in the cell.
                            saRet[iRow, iCol] = iRow.ToString() + "|" + iCol.ToString();
                        }
                    }

                    //Set the range value to the array.
                    range.set_Value(Missing.Value, saRet);
                }

                //Return control of Excel to the user.
                objApp.Visible = true;
                objApp.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                //MessageBox.Show(errorMessage, "Error");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application objApp;
            Excel._Workbook objBook;

            Excel.Workbooks objBooks;
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range range;
            try
            {
                try
                {
                    //Get a reference to the first sheet of the workbook.
                    objApp = new Excel.Application();
                    objBooks = objApp.Workbooks;
                    objBook = objBooks.Add(Missing.Value);
                    objSheets = objBook.Worksheets;
                    objSheet = (Excel._Worksheet)objSheets.get_Item(1);

                }

                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Can't find the Excel workbook.  Try clicking Button1 " +
                       "to create an Excel workbook with data before running Button2.";

                    //MessageBox.Show(errorMessage, "Missing Workbook?");

                    //You can't automate Excel if you can't find the data you created, so 
                    //leave the subroutine.
                    return;
                }

                //Get a range of data.
                range = objSheet.get_Range("A1", "E5");

                //Retrieve the data from the range.
                Object[,] saRet;
                saRet = (System.Object[,])range.get_Value(Missing.Value);

                //Determine the dimensions of the array.
                long iRows;
                long iCols;
                iRows = saRet.GetUpperBound(0);
                iCols = saRet.GetUpperBound(1);

                //Build a string that contains the data of the array.
                String valueString;
                valueString = "Array Data\n";

                for (long rowCounter = 1; rowCounter <= iRows; rowCounter++)
                {
                    for (long colCounter = 1; colCounter <= iCols; colCounter++)
                    {

                        //Write the next value into the string.
                        valueString = String.Concat(valueString,
                           saRet[rowCounter, colCounter].ToString() + ", ");
                    }

                    //Write in a new line.
                    valueString = String.Concat(valueString, "\n");
                }

                //Report the value of the array.
                //MessageBox.Show(valueString, "Array Values");
            }

            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                //MessageBox.Show(errorMessage, "Error");
            }

        }
        private void button4_Click(object sender, EventArgs e)
        {
            /**
            Read XLS or XLSX File
            anchor-read-an-xls-or-xlsx-file
            **/
            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workbook = WorkBook.Load(@"C:\Excel\BP2.xlsm");
            WorkSheet sheet = workbook.WorkSheets.First();
            WorkSheet oldSheet = workbook.GetWorkSheet("Blad9");
            //Select cells easily in Excel notation and return the calculated value
            //int cellValue = sheet["A2"].IntValue;
            // //MessageBox.Show("cellValue {0}", cellValue.ToString());

            var range = sheet["C14:D20"];
            // Read from Ranges of cells elegantly.

            foreach (var cell in range)
            {

                ////MessageBox.Show(cell.AddressString + " = " + cell.Value.ToString());

                if (cell.Value.ToString().Contains("2,000001"))
                {
                    //MessageBox.Show(cell.AddressString + " Innehåller 2,000001  ");
                    //MessageBox.Show(cell.Address + " Address ");
                    //MessageBox.Show(cell.ColumnIndex.ToString() + " ColumnIndex  ");
                    //MessageBox.Show(cell.RowIndex.ToString() + " RowIndex  ");
                    //MessageBox.Show(cell.Location + " Location ");
                    //MessageBox.Show(cell.StringValue + " StringValue ");
                    //MessageBox.Show(cell.Value + " Value ");
                }

                /*   for (int i = 0; i < range.Count(); i++)
                   {
                       //MessageBox.Show(cell.AddressString + " Innehåller G ");
                   }*/
            }

            /*  if (sheet["D1"].Value.ToString().Contains("G") )
              {
                  Console.WriteLine("Basic test passed");
              }*/

            ///Advanced Operations
            //Calculate aggregate values such as Min, Max and Sum
            //   decimal sum = sheet["A2:A10"].Sum();

            //Linq compatible
            //  decimal max = sheet["A2:A10"].Max(c => c.DecimalValue);
        }


    }
}



////MessageBox.Show(max.ToString("F2"));
////MessageBox.Show(ColumnIndex.ToString());

////MessageBox.Show(cell.AddressString + " = " + cell.Value.ToString());   if (cell.Value.ToString().Contains("2,000001")|| cell.Value.ToString().Contains("3,000002"))
//  {    }

/*
                //MessageBox.Show(cell.AddressString + " Innehåller 2,000001  ");

                //MessageBox.Show(cell.Address + " Address ");

                //MessageBox.Show(cell.ColumnIndex.ToString() + " ColumnIndex  ");

                //MessageBox.Show(cell.RowIndex.ToString() + " RowIndex  ");
                //MessageBox.Show(cell.Location + " Location ");

                //MessageBox.Show(cell.StringValue + " StringValue ");

                //MessageBox.Show(cell.Value + " Value ");
*//*
foreach (var cell in range)
{



// Read C column on cell at a time value [Opt. thickness.] and show on //MessageBox.
double CCell = sheet[$"C{cell.RowIndex }"].DoubleValue;
//MessageBox.Show(CCell.ToString() + " --->>> Värden CCell ");


// Read D column on cell at a time value [SkiktNummer] and show on //MessageBox.
string DCell = sheet[$"D{cell.RowIndex }"].ToString();
//MessageBox.Show(DCell + " --->>> Material DCell ");

// Read E column on cell at a time value [Nr] and show on //MessageBox.
double ECell = sheet[$"E{cell.RowIndex }"].DoubleValue;
//MessageBox.Show(ECell.ToString() + " --->>> right ECell ");


            string maxAdress = sheet["C14:D24"].Max(c => c.AddressString);
            int ColumnIndex = sheet["C14:D24"].Max(c => c.ColumnIndex); // shows column pos Ie D
            string MaxIndex = sheet["C14:D24"].Max(c => c.Address.ToString()); // shows last cell in range

            //string maxValue = sheet["C14:C24"].Max(c => c.DecimalValue.ToString("#,###.00")); //not wotking
            double max = sheet["C14:C24"].Max(c => c.DoubleValue);


}*/
/*  if (cell.Value.ToString().Contains("2,0") || cell.Value.ToString().Contains("4,0")|| cell.Value.ToString().Contains("6,0"))
 { }
   string rightNRCell = sheet[$"D{cell.RowIndex + 2}"].ToString();
    //MessageBox.Show(rightNRCell + " --->>> right NR Cell ");

    string rightNRCell3 = sheet[$"D{cell.RowIndex +3}"].ToString();
    //MessageBox.Show(rightNRCell3 + " --->>> right NR Cell ");
*/



/*   for (int i = 0; i < range.Count(); i++)
   {
       //MessageBox.Show(cell.AddressString + " Innehåller G ");
   }*/