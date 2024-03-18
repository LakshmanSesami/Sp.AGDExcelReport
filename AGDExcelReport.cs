using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
//using System.Data.OleDb;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;
using System.Net.Mail;
using System.Data.SqlClient;

namespace Sp.AGD.ExcelReport
{

    class AGDExcelReport
    {
        static void Main(string[] args)
        {
            
            AGDREPORT.GenerateExcelReport();
        }
    }
    class AGDREPORT
    {
        public static void GenerateExcelReport()
        {
            string dbConnectionString = ConfigurationManager.ConnectionStrings["AGDExcelReport"].ConnectionString;
            DataSet ds = new DataSet("New_DataSet");
            System.Data.DataTable dt = new System.Data.DataTable("New_DataTable");

            //Set the locale for each
            ds.Locale = System.Threading.Thread.CurrentThread.CurrentCulture;
            dt.Locale = System.Threading.Thread.CurrentThread.CurrentCulture;

            //Open a DB connection (in this example with OleDB)
            //OleDbConnection con = new OleDbConnection(dbConnectionString);
            SqlConnection con = new SqlConnection(dbConnectionString);
            con.Open();

            //Create a query and fill the data table with the data from the DB
            string sql = System.Configuration.ConfigurationManager.AppSettings["Query2"];
            SqlCommand cmd = new SqlCommand(sql, con);
            {
                cmd.CommandTimeout = 0;
            }
            //OleDbCommand cmd = new OleDbCommand(sql, con);
            SqlDataAdapter adptr = new SqlDataAdapter(cmd);
            //OleDbDataAdapter adptr = new OleDbDataAdapter();

            adptr.SelectCommand = cmd;
            adptr.Fill(dt);
            dt.TableName = "Bad Invoices";
            ds.Tables.Add(dt);

            sql = System.Configuration.ConfigurationManager.AppSettings["Query1"]; ;
             cmd = new SqlCommand(sql, con);
             adptr = new SqlDataAdapter();

            adptr.SelectCommand = cmd;
            dt = new System.Data.DataTable();
            adptr.Fill(dt);
            dt.TableName = "Good Invoices";
            //Add the table to the data set
            ds.Tables.Add(dt);            

            con.Close();

            DayOfWeek currentDay = DateTime.Now.DayOfWeek;
            int daysTillCurrentDay = currentDay - DayOfWeek.Monday;
            //DateTime currentWeekStartDate = DateTime.Now.AddDays(-daysTillCurrentDay);
            //string WeekStartDate = currentWeekStartDate.ToString("dd-MMMM-yyyy");
            DateTime lastWeekStartDate = DateTime.Now.AddDays(-daysTillCurrentDay - 7);
            string WeekStartDate = lastWeekStartDate.ToString("yyyyMMdd");
            DateTime lastWeekEndDate = DateTime.Now.AddDays(-daysTillCurrentDay - 1);
            string WeekEndDate = lastWeekEndDate.ToString("yyyyMMdd");

            string CurrDate = DateTime.Now.ToString("yyyyMMdd");
            //string EndDate = DateTime.Now.AddDays(-daysTillCurrentDay - 3).ToString("yyyyMMdd");
            //WriteDataTableToExcel(ds, "F:\\Batchjobs\\Sp.AGD.ExcelReport\\Input\\AGDExcelReport_" + CurrDate+".xlsx", "Details");
            WriteDataTableToExcel(ds, "F:\\Batchjobs\\Sp.AGD.ExcelReport\\Input\\AGD-Weekly-Report-" + WeekStartDate + "-" + WeekEndDate + ".xlsx", "Details");
            //WriteDataTableToExcel(ds, "F:\\Test\\AGD-Weekly-Report-" + WeekStartDate + "-" + WeekEndDate + ".xlsx", "Details");

            string CurrDat = DateTime.Now.ToString("dd-MMMM-yyyy");
            //DayOfWeek currentDay = DateTime.Now.DayOfWeek;
            //int daysTillCurrentDay = currentDay - DayOfWeek.Monday;
            //DateTime currentWeekStartDate = DateTime.Now.AddDays(-daysTillCurrentDay);
            //string WeekStartDate = currentWeekStartDate.ToString("dd-MMMM-yyyy");
            //DateTime lastWeekStartDate = DateTime.Now.AddDays(-daysTillCurrentDay - 7);
            string WeekStartDate1 = lastWeekStartDate.ToString("dd-MMMM-yyyy");
            //DateTime lastWeekEndDate = DateTime.Now.AddDays(-daysTillCurrentDay - 3);
            string WeekEndDate1 = lastWeekEndDate.ToString("dd-MMMM-yyyy");
            string Subject = "Weekly report of PEPPOL invoices received from " + WeekStartDate1 + " to " + WeekEndDate1;
            //var path = @"D:\\BTSFS\\AGD.ExcelReport\\Input\\"+ content + ".txt";
            string body = "Attached is the weekly report of PEPPOL invoices received from " + WeekStartDate1 + " to " + WeekEndDate1;
            //File.WriteAllText(path, text);
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient();
            String addressees = System.Configuration.ConfigurationManager.AppSettings["MailTo"];
            String[] addr = addressees.Split(',');
            foreach (string MultiEmailId in addr)
            {
                mail.To.Add(new MailAddress(MultiEmailId));
            }
            //mail.To.Add(new MailAddress(System.Configuration.ConfigurationManager.AppSettings["MailTo"]));
            mail.From = new MailAddress(System.Configuration.ConfigurationManager.AppSettings["MailFrom"]);
            mail.CC.Add(System.Configuration.ConfigurationManager.AppSettings["MailCc"]);
            mail.Subject = Subject;
            mail.IsBodyHtml = true;
            mail.Body = body;
            SmtpServer.Host = System.Configuration.ConfigurationManager.AppSettings["ServerHost"];
            SmtpServer.Port = 25;
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment("F:\\Batchjobs\\Sp.AGD.ExcelReport\\Input\\AGD-Weekly-Report-" + WeekStartDate + "-" + WeekEndDate + ".xlsx");
            mail.Attachments.Add(attachment);
            SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;

            SmtpServer.Send(mail);


            Sp.Common.Logging.LogObject objLogObject = new Sp.Common.Logging.LogObject();
            objLogObject.AppName = "Sp.AGD.ExcelReport";
            objLogObject.TransactionType = "AGD-Weekly-Report-" + WeekStartDate + "-" + WeekEndDate;
            objLogObject.EventID = "strSuccessEventID";
            objLogObject.ProcessName = "AGD.EXCEL";
            objLogObject.ReferenceNumber = "AGD-Weekly-Report-" + WeekStartDate + "-" + WeekEndDate;

            Sp.Common.Logging.NLogManager objLogMgr = new Sp.Common.Logging.NLogManager(objLogObject);

            objLogMgr.Log(Sp.Common.Logging.LogType.Info, "AGD.Excel_Success: " + "Email Sent Successfully with Attachment");


        }

        public static bool WriteDataTableToExcel(System.Data.DataSet dataSet, string saveAsLocation, string ReporType)

        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);
                ///////////////////////////////////////////////////////////////////////////////////////////
                // Workk sheet

                int count = 1;
                foreach (System.Data.DataTable datatable in dataSet.Tables)
                {
                    
                    CreateWorksheet(datatable,ReporType, excelworkBook, out excelSheet, out excelCellrange);
                    
                    if (count < dataSet.Tables.Count)
                    {
                        excelworkBook.Worksheets.Add();
                    }
                    count++;
                }
                

                ////////////////////////////////////////////////////////////////////////////////////////
                //now save the workbook and exit Excel


                excelworkBook.SaveAs(saveAsLocation); ;
                excelworkBook.Close();
                excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
             //   MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }

        private static void CreateWorksheet(System.Data.DataTable dataTable, string ReporType, Workbook excelworkBook, out Worksheet excelSheet, out Range excelCellrange)
        {
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = dataTable.TableName;           
            

            //excelSheet.Cells[1, 1] = ReporType;
            //excelSheet.Cells[1, 2] = "Date : " + DateTime.Now.ToShortDateString();

            // loop through each row and add values to our sheet
            //int rowcount = 1;

            //foreach (DataColumn datacolumn in dataTable.Columns)
            //{

            //    rowcount += 1;
            //    for (int i = 1; i <= dataTable.Columns.Count; i++)
            //    {
            //        // on the first iteration we add the column headers
            //        if (rowcount == 2)
            //        {
            //            excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            //            excelSheet.Cells.Font.Color = System.Drawing.Color.Black;
            //            //excelSheet.Cells.Font.Bold = true;

            //        }
            //        //foreach (DataRow datarow in dataTable.Rows)
            //        //{
            //        //excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
            //        //}
            //        //Range rg = excelSheet.Cells[1, 1];
            //        //rg.EntireColumn.NumberFormat = "MM/DD/YYYY";


            //        //for alternate rows
            //        //if (rowcount > 3)
            //        //{
            //        //    if (i == dataTable.Columns.Count)
            //        //    {
            //        //        if (rowcount % 2 == 0)
            //        //        {
            //        //            excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
            //        //            //FormattingExcelCells(excelCellrange, "#CCCCFF", System.Drawing.Color.Black, false);
            //        //        }

            //        //    }
            //        //}

            //    }

            //}

            //for (int i = 1; i <= dataTable.Columns.Count; i++)
            //{
            //    excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            //    excelSheet.Cells.Font.Color = System.Drawing.Color.Black;
            //    excelSheet.Cells.Font.Bold = true;
            //}

            //int rowcount = 1;
            //foreach (DataRow datarow in dataTable.Rows)
            //{

            //    rowcount += 1;
            //    for (int i = 1; i <= dataTable.Columns.Count; i++)
            //    {
            //        // on the first iteration we add the column headers
            //        //if (rowcount == 2)
            //        //{
            //        //    excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            //        //    excelSheet.Cells.Font.Color = System.Drawing.Color.Black;
            //        //}

            //        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
            //        excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

            //        excelSheet.Cells.NumberFormat = "@";
            //        //Range rg = excelSheet.Cells[1, 1];
            //        //rg.EntireColumn.NumberFormat = "MM/DD/YYYY";

            //    }
            //    Console.WriteLine(System.DateTime.Now);
            //}

            for (int Idx = 0; Idx < dataTable.Columns.Count; Idx++)
            {
                excelSheet.Range["A1"].Offset[0, Idx].Value = dataTable.Columns[Idx].ColumnName;
                //row header styles  
                excelSheet.Range["A1"].Offset[0, Idx].Font.Color = System.Drawing.Color.Black;
                excelSheet.Range["A1"].Offset[0, Idx].Font.FontStyle = "Bold";

                excelSheet.Range["A1"].Offset[0, Idx].ColumnWidth = 15;
                excelSheet.Range["A1"].Offset[0, Idx].RowHeight = 15;
                excelSheet.Range["A1"].Offset[0, Idx].WrapText = true;
                excelSheet.Range["A1"].Offset[0, Idx].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; ;
                excelSheet.Range["A1"].Offset[0, Idx].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop; ;

                Microsoft.Office.Interop.Excel.Borders borderh = excelSheet.Range["A1"].Offset[0, Idx].Borders;
                borderh.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                borderh.Weight = 2d;


            }
            for (int Idx = 0; Idx < dataTable.Rows.Count; Idx++)
            {
                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].Value =
                dataTable.Rows[Idx].ItemArray.Select(x => x.ToString()).ToArray();

                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].ColumnWidth = 15;
                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].RowHeight = 15;
                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].WrapText = true;
                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; ;
                excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop; ;

                //excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].Errors[Microsoft.Office.Interop.Excel.XlErrorChecks.xlNumberAsText].Ignore = true;

                Microsoft.Office.Interop.Excel.Borders border = excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                //excelSheet.Range["A2"].Offset[Idx].Resize[1, dataTable.Columns.Count].NumberFormat = "@";
            }

            //excelSheet.Activate();
            excelSheet.Application.ActiveWindow.SplitRow = 1;
            excelSheet.Application.ActiveWindow.FreezePanes = true;

            // now we resize the columns
            excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[dataTable.Rows.Count + 1, dataTable.Columns.Count]];
            //excelCellrange.Cells.Errors[Microsoft.Office.Interop.Excel.XlErrorChecks.xlNumberAsText].Ignore = true;

            excelSheet.Application.ErrorCheckingOptions.NumberAsText = false;
            //excelCellrange.NumberFormat = "@";
            //excelCellrange.EntireColumn.ColumnWidth = 15;
            //excelCellrange.EntireColumn.WrapText = true;
            //excelCellrange.EntireColumn.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            //excelCellrange.EntireRow.RowHeight = 15;

            //excelCellrange.EntireColumn.AutoFit();
            //excelCellrange.EntireRow.AutoFit();
            //Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            //border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            //border.Weight = 2d;


            //excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, dataTable.Columns.Count]];
            //excelCellrange.Font.Bold = true;
            //FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

            //excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, 2]];
            //excelCellrange.NumberFormat = "MM/DD/YYYY HH:mm:ss AM";
            //excelCellrange.EntireColumn.ColumnWidth = 15;
            //excelCellrange.EntireRow.RowHeight = 15;
            //excelCellrange.EntireColumn.WrapText = true;
            //excelCellrange.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
        }

        public static void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

    }
}
