using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //string filePath = @"D:\DOTNET\WebForm\20170619130221409-Test-schedule.xlsx";
        string filePath = @"D:\DOTNET\WebForm\MyFile.xls";

        DataTable dtExcelData = GenerateExcelData(filePath);
    }

    public DataTable GenerateExcelData(string Filepath)
    {
        DataTable dtValues = null;
        IExcelDataReader excelReader = null;
        try
        {
            FileStream stream = File.Open(Filepath, FileMode.Open, FileAccess.Read);
            if (Path.GetExtension(Filepath) == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (Path.GetExtension(Filepath) == ".xlsx")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            if (excelReader != null)
            {
                DataSet ds = excelReader.AsDataSet();
                if (ds != null && ds.Tables.Count > 0)
                {
                    dtValues = ds.Tables[0];
                }
            }
        }
        // need to catch possible exceptions
        catch (Exception ex)
        {
        }
        finally
        {
            excelReader.Close();
        }
        return dtValues;
    }
}