using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace RevitCommon
{
    public class Excel
    {
        /// <summary>
        /// The ExcelWrite will write lists of data to a selected Excel file
        /// </summary>
        /// <param name="filePath">Path to the XLS/XLSX file.</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="data">Data to be written</param>
        /// <param name="transpose">Data should be in row format.  Transpose takes columnar data and turns it into row data.</param>
        /// <returns></returns>
        public static bool Write(string filePath, string worksheetName, List<List<string>> data, bool transpose)
        {
            bool write = Write(filePath, worksheetName, data, transpose, new List<bool> { false });
            return write;
        }

        /// <summary>
        /// The ExcelWrite will write lists of data to a selected Excel file
        /// </summary>
        /// <param name="filePath">Path to the XLS/XLSX file.</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="data">Data to be written</param>
        /// <param name="transpose">Data should be in row format.  Transpose takes columnar data and turns it into row data.</param>
        /// <param name="forceText">A list of bools to specify whether a column should be forced to be text or not.
        /// If a single bool is specified in the list, then that will be used for all columns.</param>
        /// <returns></returns>
        public static bool Write(string filePath, string worksheetName, List<List<string>> data, bool transpose, List<bool> forceText)
        {
            try
            {
                List<List<string>> writeData;
                if (transpose)
                    writeData = Transpose(data);
                else
                    writeData = data;

                // Open the Excel file and get the worksheet
                MSExcel.Application excelApp = null;
                bool alreadyOpen = false;
                try
                {
                    excelApp = (MSExcel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null && excelApp.Visible)
                        alreadyOpen = true;
                }
                catch { }
                if (excelApp == null)
                    excelApp = new MSExcel.Application();
                excelApp.Visible = true;

                MSExcel.Workbook workbook = null;
                bool existingWB = false;

                if (filePath == null)
                {
                    workbook = excelApp.Workbooks.Add();
                    existingWB = false;
                }
                else
                {
                    // See if the workbook is open already
                    foreach (MSExcel.Workbook wb in excelApp.Workbooks)
                    {
                        if (wb.FullName.ToLower() == filePath.ToLower())
                        {
                            workbook = wb;
                            existingWB = true;
                            break;
                        }
                    }

                    // Try opening the file...
                    try
                    {
                        if (workbook == null)
                            workbook = excelApp.Workbooks.Open(filePath);
                        existingWB = true;
                    }
                    // Alright, just create a new file.
                    catch
                    {
                        workbook = excelApp.Workbooks.Add();
                        existingWB = false;
                        //alreadyOpen = false;
                    }
                }

                MSExcel.Worksheet worksheet = null;
                foreach (MSExcel.Worksheet ws in workbook.Sheets)
                {
                    if (ws.Name.ToLower() == worksheetName.ToLower())
                    {
                        worksheet = ws;
                        break;
                    }
                }

                if (worksheet == null)
                {
                    try
                    {
                        worksheet = workbook.Worksheets.Add();
                        worksheet.Name = worksheetName;
                    }
                    catch { }
                }
                if (worksheet == null)
                {
                    System.Windows.Forms.MessageBox.Show("Error:\n" + "Worksheet is still null");
                }

                // Write all of the data to the excel file
                for (int i = 0; i < writeData.Count; i++)
                {
                    for (int j = 0; j < writeData[i].Count; j++)
                    {
                        bool forcetext = false;
                        if (forceText.Count > j)
                            forcetext = forceText[j];
                        else
                            forcetext = forceText.Last();
                        try
                        {
                            MSExcel.Range cell = worksheet.Cells[i + 1, j + 1];
                            if (forcetext)
                                cell.NumberFormat = "@";
                            cell.Value2 = writeData[i][j];
                        }
                        catch
                        {
                            System.Windows.Forms.MessageBox.Show("Error");
                        }
                    }
                }

                // Close the file if necessary
                if (!alreadyOpen)
                {
                    if (existingWB)
                    {
                        workbook.Save();
                        workbook.Close();
                        excelApp.Quit();
                    }
                    else if (filePath != null && !existingWB)
                    {
                        //System.Windows.Forms.MessageBox.Show("AlreadyOpen: " + alreadyOpen.ToString() + "\nExisting WB: " + existingWB + "\nFilePath:" + filePath);
                        workbook.SaveAs(filePath);
                        workbook.Close();
                        excelApp.Quit();
                    }
                }
                else if (filePath != null && !existingWB)
                {
                    //System.Windows.Forms.MessageBox.Show("AlreadyOpen: " + alreadyOpen.ToString() + "\nExisting WB: " + existingWB + "\nFilePath:" + filePath);
                    workbook.SaveAs(filePath);
                    workbook.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Excel Write Error\n" + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// The ExcelRead method returns a multidimensional list of data organized by rows.  The first row will
        /// be read as the headers of the file, with all other rows making up the data.
        /// </summary>
        /// <param name="filePath">Path to the XLS/XLSX file.</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <returns>2D list of strings that are organized as rows of data.</returns>
        public static List<List<string>> Read(string filePath, string worksheetName)
        {

            List<List<string>> data = new List<List<string>>();

            // Open the Excel file and get the worksheet
            MSExcel.Application excelApp = null;
            bool closeExcel = true;
            try
            {
                excelApp = (MSExcel.Application)Marshal.GetActiveObject("Excel.Application");
                closeExcel = false;
            }
            catch { }
            if (excelApp == null)
                excelApp = new MSExcel.Application();


            MSExcel.Workbook workbook = null;
            bool closeWorkbook = true;
            // See if the workbook is open already
            foreach (MSExcel.Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.ToLower() == filePath.ToLower())
                {
                    workbook = wb;
                    closeWorkbook = false;
                    break;
                }
            }
            if (workbook == null)
            {
                var workbooks = excelApp.Workbooks;
                workbook = excelApp.Workbooks.Open(filePath);
                Marshal.ReleaseComObject(workbooks);
            }

            MSExcel.Worksheet worksheet = null;
            MSExcel.Sheets worksheets = workbook.Sheets;
            foreach (MSExcel.Worksheet ws in worksheets)
            {
                Console.WriteLine(worksheetName);
                if (ws.Name.ToLower() == worksheetName.ToLower())
                {
                    worksheet = ws;
                    break;
                }
            }
            

            // Read through the data in the excel file and add it to our data variable.
            if (null != worksheet)
            {
                int rowCount = 0;
                int colCount = 0;

                // Get the total used range of the selected worksheet
                MSExcel.Range usedRange = worksheet.UsedRange;
                colCount = usedRange.Columns.Count;
                rowCount = usedRange.Rows.Count;

                // Read the first row as headers
                List<string> headers = new List<string>();
                for (int i = 1; i <= colCount; i++)
                {
                    try
                    {
                        MSExcel.Range headerCell = usedRange.Cells[1, i];
                        headers.Add(headerCell.Text);
                        Marshal.ReleaseComObject(headerCell);
                    }
                    catch { }
                }
                data.Add(headers);

                // Iterate through the rows and add the data
                for (int i = 2; i <= rowCount; i++)
                {
                    List<string> rowData = new List<string>();
                    for (int j = 1; j <= colCount; j++)
                    {
                        MSExcel.Range cell = usedRange.Cells[i, j];
                        try
                        {
                            rowData.Add(cell.Text);
                        }
                        catch
                        {
                            rowData.Add(string.Empty);
                        }
                        Marshal.ReleaseComObject(cell);
                    }
                    data.Add(rowData);
                }
                Marshal.ReleaseComObject(usedRange);
            }
            

            // Close the worksheet
            //worksheet = null;
            //worksheets = null;
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(worksheets);
            // Close the workbook if it was opened via this command
            if (closeWorkbook)
            {
                workbook.Close(false, filePath, Missing.Value);
                Marshal.ReleaseComObject(workbook);
            }
            else
                Marshal.ReleaseComObject(workbook);

            if (closeExcel)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            else
                Marshal.ReleaseComObject(excelApp);


            //  Send that sweet data home.
            return data;
        }

        /// <summary>
        /// Get the names of all worksheets in the selected Excel file.
        /// </summary>
        /// <param name="filePath">Path to an excel file</param>
        /// <returns></returns>
        public static List<string> GetWorksheetNames(string filePath)
        {
            List<string> worksheets = new List<string>();

            // Open the Excel file and get the worksheet
            MSExcel.Application excelApp = null;
            bool closeExcel = true;
            try
            {
                excelApp = (MSExcel.Application)Marshal.GetActiveObject("Excel.Application");
                closeExcel = false;
            }
            catch { }
            if (excelApp == null)
                excelApp = new MSExcel.Application();


            MSExcel.Workbook workbook = null;
            bool closeWorkbook = true;
            // See if the workbook is open already
            foreach (MSExcel.Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.ToLower() == filePath.ToLower())
                {
                    workbook = wb;
                    closeWorkbook = false;
                    break;
                }
            }
            if (workbook == null)
            {
                var workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(filePath);
                Marshal.ReleaseComObject(workbooks);
            }

            // Get the worksheet names
            foreach (MSExcel.Worksheet ws in workbook.Sheets)
            {
                worksheets.Add(ws.Name);
                Marshal.ReleaseComObject(ws);
            }

            if (closeWorkbook)
            {
                workbook.Close(false, filePath, Missing.Value);
                Marshal.ReleaseComObject(workbook);
            }
            
            if(closeExcel)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return worksheets;
        }

        /// <summary>
        /// Transpose the incoming data from being organized in
        /// columns to being organized into rows or vice versa.
        /// </summary>
        /// <param name="incomingData">Data to transpose</param>
        /// <returns></returns>
        internal static List<List<string>> Transpose(List<List<string>> incomingData)
        {
            List<List<string>> data = new List<List<string>>();

            for (int j = 0; j < incomingData[0].Count; j++)
            {
                List<string> rowData = new List<string>();
                for (int i = 0; i < incomingData.Count; i++)
                {
                    try
                    {
                        rowData.Add(incomingData[i][j]);
                    }
                    catch
                    {
                        rowData.Add("ERROR");
                    }
                }
                data.Add(rowData);
            }

            return data;
        }
    }
}
