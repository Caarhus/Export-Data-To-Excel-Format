using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExportDataTableToExcel.Models;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;

namespace ExportDataTableToExcelInMVC4.Controllers
{
    public class ExportDataController : Controller
    {
        public ActionResult Index()
        {
            String constring = ConfigurationManager.ConnectionStrings["RConnection"].ConnectionString;
            SqlConnection con = new SqlConnection(constring);
            string query = "select * From Employee";
            DataTable dt = new DataTable();
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter(query, con);
            da.Fill(dt);
            con.Close();
            IList<ExportDataTableToExcelModel> model = new List<ExportDataTableToExcelModel>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                model.Add(new ExportDataTableToExcelModel()
                {
                    Id = Convert.ToInt32(dt.Rows[i]["Id"]),
                    Name = dt.Rows[i]["Name"].ToString(),
                    Email = dt.Rows[i]["Email"].ToString(),
                    Country = dt.Rows[i]["Country"].ToString(),
                });
            }
            return View(model);
        }

        public ActionResult ExportData()
        {
            String constring = ConfigurationManager.ConnectionStrings["RConnection"].ConnectionString;
            SqlConnection con = new SqlConnection(constring);
            string query = "select * From Employee";
            DataTable dt = new DataTable
            {
                TableName = "Employee"
            };
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter(query, con);
            da.Fill(dt);
            con.Close();

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename= EmployeeReport.xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
            return RedirectToAction("Index", "ExportData");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}