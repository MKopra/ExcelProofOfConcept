using System;
using ExcelProofOfConcept;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using static System.Runtime.InteropServices.JavaScript.JSType;


internal class Program
{
    public static void Main(string[] args)
    {

        var companyInformation = new CompanyInformation("");


        //assign excel doc we want to read to a var
        using (var package = new ExcelPackage(new FileInfo("NOV 21 service 2-82 plus CAL.xlsx"))) //
        {
            var firstSheet = package.Workbook.Worksheets["Sheet1"]; //selects sheet
            Console.WriteLine("Enter your Company (Ex: A,B,C,HHC,FSC)\n");
            var userInput = Console.ReadLine();
            Console.WriteLine("All services or just overdue? (Ex: all, all abreviated, or overdue)\n");
            var overdueOrAll = Console.ReadLine();
            Console.WriteLine("What type of date would you like (Ex: early, planned, or late)\n");
            var userDate = Console.ReadLine();
            var outputSheet = package.Workbook.Worksheets.Add("Sheet2");
            outputSheet.TabColor = System.Drawing.Color.Black;
            outputSheet.DefaultRowHeight = 12;
            // Setting the properties
            // of the first row
            outputSheet.Row(1).Height = 20;
            outputSheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            outputSheet.Row(1).Style.Font.Bold = true;
            outputSheet.Row(1).Style.Border.BorderAround(ExcelBorderStyle.Hair);
            //set column widths for results
            outputSheet.Column(1).Width = 20;
            outputSheet.Column(2).Width = 20;
            outputSheet.Column(3).Width = 20;
            outputSheet.Column(4).Width = 40;
            outputSheet.Column(5).Width = 20;
            outputSheet.Column(6).Width = 20;
            outputSheet.Column(7).Width = 20;
            outputSheet.Cells["A1"].Value = "UIC";
            outputSheet.Cells["B1"].Value = "Admin Number";
            outputSheet.Cells["C1"].Value = "Model Number";
            outputSheet.Cells["D1"].Value = "Description";
            if (userDate is "early")
            {
                outputSheet.Cells["E1"].Value = "Early Date";
                outputSheet.Cells["F1"].Value = "Days Until Early Date";
            }
            if (userDate is "planned")
            {
                outputSheet.Cells["E1"].Value = "Planned Date";
                outputSheet.Cells["F1"].Value = "Days Until Planned Date";
            }
            if (userDate is "late")
            {
                outputSheet.Cells["E1"].Value = "Late Date";
                outputSheet.Cells["F1"].Value = "Days Until Late Date";
            }
            package.Save();
            var sheet1lastRow = firstSheet.Dimension.End.Row;
            var lastRowOutputSheet = outputSheet.Row(sheet1lastRow);

            var dayToday = DateTime.Now.ToString("MM/dd/yyyy");

            //Console.WriteLine(dayToday);
            if (userInput is "A" && overdueOrAll is "all" or "all abbreviated")
            {
                var start = firstSheet.Dimension.Start;
                var end = firstSheet.Dimension.End;
                //var startOutput = outputSheet.Dimension.Start;
                //var endOutput = lastRowOutputSheet;
                for (int row = start.Row; row <= end.Row; row++)
                {
                    ///////char fifthChar = firstSheet.Cells[row, 1].[4];
                    if (firstSheet.Cells[row, 1].Value.Equals("WAD8A0"))
                    {
                        //for (int rowOutput = startOutput.Row; rowOutput <= endOutput.Row; rowOutput++) //when I tried this it worked but pasted the same values across 4000 rows
                        firstSheet.Cells[row, 1].Copy(outputSheet.Cells[row, 1]); //UIC
                        firstSheet.Cells[row, 4].Copy(outputSheet.Cells[row, 2]); //adminNum
                        firstSheet.Cells[row, 5].Copy(outputSheet.Cells[row, 3]); //modelNum
                        firstSheet.Cells[row, 7].Copy(outputSheet.Cells[row, 4]); //Description
                        if (userDate is "early")
                        {
                            int dateRow = 10;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (userDate is "planned")
                        {
                            int dateRow = 11;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (userDate is "late")
                        {
                            int dateRow = 12;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (firstSheet.Cells[row, 2].Value.Equals("OVERDUE"))
                        {
                            outputSheet.Cells[row, 7].Value = "*OVERDUE*";
                        }
                        var startOutput = outputSheet.Dimension.Start;
                        var endOutput = lastRowOutputSheet;
                        for (int rowOutput = startOutput.Row; rowOutput <= endOutput.Row; rowOutput++)
                        {
                            var rowCell = outputSheet.Cells[rowOutput, 1].Value;
                            if (rowCell is null)
                            {
                                outputSheet.DeleteRow(rowOutput);
                            }
                        }
                    }
                    package.Save();

                }
            }
            if (userInput is "A" && overdueOrAll is "overdue")
            {
                var start = firstSheet.Dimension.Start;
                var end = firstSheet.Dimension.End;
                //var startOutput = outputSheet.Dimension.Start;
                //var endOutput = lastRowOutputSheet;
                for (int row = start.Row; row <= end.Row; row++)
                {
                    if ((firstSheet.Cells[row, 1].Value.Equals("WAD8A0")) && (firstSheet.Cells[row, 2].Value.Equals("OVERDUE")))
                    {
                        //for (int rowOutput = startOutput.Row; rowOutput <= endOutput.Row; rowOutput++) //when I tried this it worked but pasted the same values across 4000 rows
                        firstSheet.Cells[row, 1].Copy(outputSheet.Cells[row, 1]); //UIC
                        firstSheet.Cells[row, 4].Copy(outputSheet.Cells[row, 2]); //adminNum
                        firstSheet.Cells[row, 5].Copy(outputSheet.Cells[row, 3]); //modelNum
                        firstSheet.Cells[row, 7].Copy(outputSheet.Cells[row, 4]); //Description
                        if (userDate is "early")
                        {
                            int dateRow = 10;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (userDate is "planned")
                        {
                            int dateRow = 11;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (userDate is "late")
                        {
                            int dateRow = 12;
                            firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                        }
                        if (firstSheet.Cells[row, 2].Value.Equals("OVERDUE"))
                        {
                            outputSheet.Cells[row, 7].Value = "*OVERDUE*";
                        }
                        var startOutput = outputSheet.Dimension.Start;
                        var endOutput = lastRowOutputSheet;
                        for (int rowOutput = startOutput.Row; rowOutput <= endOutput.Row; rowOutput++)
                        {
                            var rowCell = outputSheet.Cells[rowOutput, 1].Value;
                            if (rowCell is null)
                            {
                                outputSheet.DeleteRow(rowOutput);
                            }
                        }
                    }
                    package.Save();

                }
            }



            var startCol6 = outputSheet.Row(2);
            var endCol6 = outputSheet.Dimension.End;
            //var startOutput = outputSheet.Dimension.Start;
            //var endOutput = lastRowOutputSheet;
            for (int row = startCol6.Row; row <= endCol6.Row; row++)
            {
                if (outputSheet.Cells[row, 5].Value is not null)

                {
                    var cellDate = outputSheet.Cells[row, 5].Value.ToString();
                    DateTime startTime = DateTime.Parse(cellDate);
                    DateTime dt = DateTime.Parse(dayToday);
                    TimeSpan t = startTime - dt;
                    outputSheet.Cells[row, 6].Value = (int)t.TotalDays;
                }
                /*if (overdueOrAll is "all abbreviated" &&(outputSheet.Cells[row, 4].Value.Equals(outputSheet.Cells[row + 1, 4].Value) && (outputSheet.Cells[row, 5].Value.Equals(outputSheet.Cells[row + 1, 5].Value))))
                {
                    //var firstRow = firstSheet.Cells[row, 7].Value;
                    var count = firstSheet.Cells[row + 1, 7].Count(c => c.Text == ("MASK SYSTEM CHEMICAL"));
                    outputSheet.Cells[row, 8].Value = count;
                }*/
            }
            package.Save();

            var startCount = outputSheet.Row(2);
            var endCount = outputSheet.Dimension.End;
            //var startOutput = outputSheet.Dimension.Start;
            //var endOutput = lastRowOutputSheet;
            for (int row = startCount.Row; row <= endCount.Row; row++)
            {
                if (overdueOrAll is "all abbreviated" )//&& (outputSheet.Cells[row, 4].Value.Equals(outputSheet.Cells[row + 1, 4].Value)))
                {
                    //var firstRow = firstSheet.Cells[row, 7].Value;
                    var count = outputSheet.Cells[row + 1, 4].Count(c => c.Text == ("MASK SYSTEM CHEMICL"));
                    outputSheet.Cells[row, 8].Value = count;
                }
                else
                {
                    outputSheet.Cells[row, 8].Value = outputSheet.Cells[row, 1].Value;
                } //if it fails the text column 8 is uic
            }
            package.Save();


            //put values from excel into local data object
            string companyName = companyInformation.CompanyName;
            firstSheet.Column(1);


            //make a new worksheet that has # of pages
            ExcelPackage excelPackage = new ExcelPackage();
            /*
            var worksheet = excelPackage.Workbook.Worksheets.Add("Data");

            int counter = 1;

            try
            {
                worksheet.Cells.AutoFitColumns();
            }
            catch (Exception) { }


            //assign the UIC column to a var named UIC
            var UIC = "WAD8A0";
            string A = "WAD8A0";
            string B = "WAD8B0";
            string C = "WAD8C0";
            string HHC = "WAD8T0";
            string FSC = "WH0KH0";

            //make a switch statement
            switch (UIC)
            {
                case A: //this takes all info after row 
                    break;

            }
            */
            // print company names in new excel doc
        }
    }
}
