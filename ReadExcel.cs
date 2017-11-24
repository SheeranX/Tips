using NPOI.HSSF.UserModel;
/*Added this class for Data Upload CR - Nov 16th 2017*/
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for ExcelUtility
/// </summary>
public class ExcelUtility
{

    public static DataTable GetExcel(string filePath)
    {
        DataTable dtExcel = new DataTable();
        try
        {
            IWorkbook workbook;

             

            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                string extension = filePath.Split('.')[1];
                if (extension == "xlsx")
                    workbook = new XSSFWorkbook(stream);
                else
                    workbook = new HSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
            dtExcel.TableName = sheet.SheetName;

            // write header row
            IRow headerRow = sheet.GetRow(0);
            

           // string title1 = headerRow.GetCell(0).StringCellValue;
            string title1 = headerRow.GetCell(1).StringCellValue;
            string title2 = headerRow.GetCell(2).StringCellValue;

            //check column name 
            if (title1.Equals(Constants.KEY_FIGURE_NAME) && title2.Equals(Constants.KF_DESCRIPTION))
            {
                int countCloumns = 0;
                foreach (ICell headerCell in headerRow)
                {
                    if (headerCell.StringCellValue.Trim()!= string.Empty)
                    {
                        dtExcel.Columns.Add(headerCell.ToString());
                        countCloumns++;
                    }
                }
                // write the rest
                int rowIndex = 0;
                foreach (IRow row in sheet)
                {
                    // skip header row
                    if (rowIndex++ == 0) continue;
                    DataRow dataRow = dtExcel.NewRow();
                    object[] rowArray = row.Cells.Select(c => c.ToString()).Take(countCloumns).ToArray();
                    dataRow.ItemArray = rowArray;
                    //ignore the empty
                    if (string.Join("", rowArray).Trim() == string.Empty)
                        break;
                    dtExcel.Rows.Add(dataRow);
                }
                return dtExcel;
            }
        }
        catch (Exception)
        {
            dtExcel.Dispose();
            File.Delete(filePath);
        }

        return dtExcel;
    }
}
