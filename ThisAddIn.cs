using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Data.SqlClient;

namespace SerialNumberRetrieval
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        /*public void runDataRetreival()
        {// SELECT CMMTTEXT WHERE SOPNUMBE='I505904'
            SerialNumberTables tables = new SerialNumberTables();
            //DataRow[] result = tables.SOP10202.Select("SOPNUMBE='I505904'");

            EnumerableRowCollection cmmttext = from result in tables.SOP10202.AsEnumerable()
                           where result.Field<String>("SOPNUMBE") == "I505904"
                           select result;

            foreach (DataRow row in cmmttext)
            {
                Console.WriteLine(row.Field<String>("CMMTTEXT"));
            }
        }*/

        private String DATASTRING = "Data Source=METRO-GP1;Integrated Security=SSPI;Initial Catalog=METRO";
        private String SELECTCOLUMNS = @"SELECT
                                            dbo.SOP30300.ShipToName,
                                            dbo.SOP30300.ACTLSHIP,
                                            dbo.SOP30300.SOPNUMBE,
                                            dbo.SOP30300.ITEMNMBR,
                                            dbo.SOP30300.QUANTITY,
                                            dbo.SOP30300.UNITCOST,
                                            dbo.SOP10202.CMMTTEXT ";
        private String FROMTABLES = "FROM dbo.SOP30300 JOIN dbo.SOP10202 ";
        private String COLUMNCORRELATION = "AND dbo.SOP30300.LNITMSEQ = dbo.SOP10202.LNITMSEQ ";
        private int NUMCOLUMNS = 7;
        private int SHIPTOCOL = 0;
        private int ACTLSHIP = 1;
        private int SOPNUMBE = 2;
        private int ITEMNMBR = 3;
        private int QUANTITY = 4;
        private int UNITCOST = 5;
        private int CMMTTEXT = 6;

        private class Row
        {
            private int SHIPTOxCOL = 1;
            private int ACTLSHIPxCOL = 2;
            private int SOPNUMBExCOL = 3;
            private int ITEMNMBRxCOL = 4;
            private int SERIALNUMxCOL = 5;
            private int QUANTITYxCOL = 6;
            private int UNITCOSTxCOL = 7;

            public DateTime invoiceDate { get; }
            public int quantity { get; }
            public int unitCost { get; }
            public string shipTo { get; }
            public string sopNumber { get; }
            public string itemNumber { get; }
            public string comment { get; }
            public string[] serialNumbers { get; }

            public Row(object reseller, object date, object invoiceNumber, object partNumber, object qty, object cost, object commentText)
            {
                /*shipTo = reseller.ToString();
                invoiceDate = (DateTime)date;
                sopNumber = invoiceNumber.ToString();
                itemNumber = partNumber.ToString();
                quantity = (int)qty;
                unitCost = (int)cost;
                comment = commentText.ToString();
                serialNumbers = commentText.ToString().Split(',');*/

                shipTo = reseller.ToString().Trim();
                invoiceDate = (DateTime)date;
                sopNumber = invoiceNumber.ToString().Trim();
                itemNumber = partNumber.ToString().Trim();
                quantity = Convert.ToInt32(qty);
                unitCost = Convert.ToInt32(cost);
                comment = commentText.ToString().Trim();
                serialNumbers = commentText.ToString().Split(',');
            }

            public void parseRow(Excel.Worksheet ws)
            {
                int rowNumber = ws.UsedRange.Row + ws.UsedRange.Rows.Count;

                if (rowNumber == 2) rowNumber = 1;

                for (int i = 0; i < serialNumbers.Length - 1; i++)
                {
                    ws.Cells[rowNumber, SHIPTOxCOL].Value = shipTo;
                    ws.Cells[rowNumber, ACTLSHIPxCOL].Value = invoiceDate;
                    ws.Cells[rowNumber, SOPNUMBExCOL].Value = sopNumber;
                    ws.Cells[rowNumber, ITEMNMBRxCOL].Value = itemNumber;
                    ws.Cells[rowNumber, SERIALNUMxCOL].Value = serialNumbers[i];
                    ws.Cells[rowNumber, QUANTITYxCOL].Value = 1;
                    ws.Cells[rowNumber, UNITCOSTxCOL].Value = unitCost;

                    rowNumber++;
                }
                //ws.Cells[rowNumber-1, 1].Value = "Last Row";
                System.Diagnostics.Debug.WriteLine("Row Number: {0}", rowNumber);
            }
        }

        private String buildQuery(String invoiceNumber)
        {
            String query = SELECTCOLUMNS + FROMTABLES;

            query += "ON dbo.SOP30300.SOPNUMBE = '" + invoiceNumber + "' AND dbo.SOP10202.SOPNUMBE ='" + invoiceNumber + "' ";
            query += COLUMNCORRELATION;

            return query;
        }

        public void runDataRetreival(string invoiceNumber)
        {
            Row thisRow;

            Excel.Workbook thisWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet thisWorksheet = thisWorkbook.Worksheets[1];
            /*int rowIndex = 1;
            int colIndex = 1;*/

            thisWorksheet.Cells[1, 2].EntireColumn.NumberFormat = "MM/DD/YYYY";
            thisWorksheet.Cells[1, 7].EntireColumn.NumberFormat = "$0.00";
            thisWorksheet.Cells[1, 5].EntireColumn.NumberFormat = "@";


            SqlConnection dbConnection = new SqlConnection(DATASTRING);
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            cmd.CommandText = buildQuery(invoiceNumber);
            cmd.CommandType = CommandType.Text;
            cmd.Connection = dbConnection;

            dbConnection.Open();

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    try
                    {
                        thisRow = new Row(
                            reader.GetValue(SHIPTOCOL), reader.GetValue(ACTLSHIP),
                            reader.GetValue(SOPNUMBE), reader.GetValue(ITEMNMBR),
                            reader.GetValue(QUANTITY), reader.GetValue(UNITCOST),
                            reader.GetValue(CMMTTEXT));

                        if (thisRow.quantity != thisRow.serialNumbers.Length - 1)
                        {
                            System.Diagnostics.Debug.WriteLine("Quantity does not match the number of serial numbers");
                        }

                        thisRow.parseRow(thisWorksheet);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.ToString());
                        throw;
                    }
                    /*thisRow = new Row(
                        reader.GetValue(SHIPTOCOL), reader.GetValue(ACTLSHIP),
                        reader.GetValue(SOPNUMBE), reader.GetValue(ITEMNMBR),
                        reader.GetValue(QUANTITY), reader.GetValue(UNITCOST),
                        reader.GetValue(CMMTTEXT));

                    if(thisRow.quantity != thisRow.serialNumbers.Length-1)
                    {
                        System.Diagnostics.Debug.WriteLine("Quantity does not match the number of serial numbers");
                    }

                    thisRow.parseRow(thisWorksheet);*/

                    /*for (int i = 0; i < NUMCOLUMNS; i++)
                    {
                        thisWorksheet.Cells[rowIndex, i + colIndex].Value = reader.GetValue(i);
                    }
                    rowIndex++;*/
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("No Rows Found");
            }
            reader.Close();
        }
    }
}
