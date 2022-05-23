using Aspose.Cells;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace MordanizerApplication.Controllers.Mordanizer
{
    public class MordanizerController : Controller
    {
        // GET: Mordanizer
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ImportFromExcel(HttpPostedFileBase postedFile)
        {
            if (ModelState.IsValid)
            {
                if (postedFile != null && postedFile.ContentLength > (1024 * 1024 * 50))  // 50MB limit  
                {
                    ModelState.AddModelError("postedFile", "Your file is to large. Maximum size allowed is 50MB !");
                }

                else
                {
                    string filePath = string.Empty;
                    string path = Server.MapPath("~/Uploads/");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    filePath = path + Path.GetFileName(postedFile.FileName);
                    string extension = Path.GetExtension(postedFile.FileName);
                    postedFile.SaveAs(filePath);

                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls": //For Excel 97-03.  
                            conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                            break;
                        case ".xlsx": //For Excel 07 and above.  
                            conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                            break;
                    }

                    try
                    {
                        DataTable dt = new DataTable();
                        using (var package = new ExcelPackage(postedFile.InputStream))
                        {
                            
                            Workbook workbook = new Workbook(postedFile.InputStream);
                            Worksheet worksheet = workbook.Worksheets[0];
                            //worksheet
                            dt = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1, true);
                       
                        }

                        var script = BuildCreateTableScript(dt, Path.GetFileName(postedFile.FileName));
                        //string[] columnNames = dt.Columns.Cast<DataColumn>()
                        //                        .Select(x => x.ColumnName)
                        //                        .ToArray();
                        //GridView dataGridView1 = new GridView();
                        //dataGridView1.DataSource = columnNames;
                        //dataGridView1.DataBind();

                        //conString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                        //using (SqlConnection con = new SqlConnection(conString))
                        //{
                        //    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        //    {
                        //        //Set the database table name.  
                        //        sqlBulkCopy.DestinationTableName = "InsuranceCertificate";
                        //        con.Open();
                        //        sqlBulkCopy.WriteToServer(dt);
                        //        con.Close();
                        //        return Json("File uploaded successfully");
                        //    }
                        //}
                    }

                    //catch (Exception ex)  
                    //{  
                    //    throw ex;  
                    //}  
                    catch (Exception e)
                    {
                        return Json("error" + e.Message);
                    }
                    //return RedirectToAction("Index");  
                }
            }
            //return View(postedFile);  
            return Json("no files were selected !");
        }
        public static string BuildCreateTableScript(DataTable Table,string TableName)
        {
            //if (!Header.IsValidDatatable(Table, IgnoreZeroRows: true))
            //    return string.Empty;

            StringBuilder result = new StringBuilder();
            result.AppendFormat("CREATE TABLE [{1}] ({0}   ", Environment.NewLine, TableName);

            result.AppendFormat("[{0}] {1} {2} {3} {4}",
                    "ID", // 0
                    "[int]", // 1
                    "IDENTITY(1,1)",//2
                    "NOT NULL", // 3
                    Environment.NewLine // 4
                    );
            result.Append("   ,");
            bool FirstTime = true;
            foreach (DataColumn column in Table.Columns.OfType<DataColumn>())
            {
                if (FirstTime) FirstTime = false;
                else
                    result.Append("   ,");

                result.AppendFormat("[{0}] {1} {2} {3}",
                    column.ColumnName, // 0
                    GetSQLTypeAsString(column.DataType), // 1
                    column.AllowDBNull ? "NULL" : "NOT NULL", // 2
                    Environment.NewLine // 3
                );
            }
            result.AppendFormat(") ON [PRIMARY]{0}", Environment.NewLine);
            //result.AppendFormat(") ON [PRIMARY]{0}GO{0}{0}", Environment.NewLine);


            // Build an ALTER TABLE script that adds keys to a table that already exists.
            if (Table.PrimaryKey.Length > 0)
                result.Append(BuildKeysScript(Table));

            return result.ToString();
        }

        /// <summary>
        /// Builds an ALTER TABLE script that adds a primary or composite key to a table that already exists.
        /// </summary>
        private static string BuildKeysScript(DataTable Table)
        {
            // Already checked by public method CreateTable. Un-comment if making the method public
            // if (Helper.IsValidDatatable(Table, IgnoreZeroRows: true)) return string.Empty;
            if (Table.PrimaryKey.Length < 1) return string.Empty;

            StringBuilder result = new StringBuilder();

            if (Table.PrimaryKey.Length == 1)
                result.AppendFormat("ALTER TABLE {1}{0}   ADD PRIMARY KEY ({2}){0}GO{0}{0}", Environment.NewLine, Table.TableName, Table.PrimaryKey[0].ColumnName);
            else
            {
                List<string> compositeKeys = Table.PrimaryKey.OfType<DataColumn>().Select(dc => dc.ColumnName).ToList();
                string keyName = compositeKeys.Aggregate((a, b) => a + b);
                string keys = compositeKeys.Aggregate((a, b) => string.Format("{0}, {1}", a, b));
                result.AppendFormat("ALTER TABLE {1}{0}ADD CONSTRAINT pk_{3} PRIMARY KEY ({2}){0}GO{0}{0}", Environment.NewLine, Table.TableName, keys, keyName);
            }

            return result.ToString();
        }

        /// <summary>
        /// Returns the SQL data type equivalent, as a string for use in SQL script generation methods.
        /// </summary>
        private static string GetSQLTypeAsString(Type DataType)
        {
            switch (DataType.Name)
            {
                case "Boolean": return "[bit]";
                case "Char": return "[char]";
                case "SByte": return "[tinyint]";
                case "Int16": return "[smallint]";
                case "Int32": return "[int]";
                case "Int64": return "[bigint]";
                case "Byte": return "[tinyint] UNSIGNED";
                case "UInt16": return "[smallint] UNSIGNED";
                case "UInt32": return "[int] UNSIGNED";
                case "UInt64": return "[bigint] UNSIGNED";
                case "Single": return "[float]";
                case "Double": return "[double]";
                case "Decimal": return "[decimal]";
                case "DateTime": return "[datetime]";
                case "Guid": return "[uniqueidentifier]";
                case "Object": return "[variant]";
                case "String": return "[nvarchar](250)";
                default: return "[nvarchar](MAX)";
            }
        }
    }
}