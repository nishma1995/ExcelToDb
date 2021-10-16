using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelToDb
{
    public partial class View : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetStudent();
        }
        public void GetStudent()   //function to get employee list
        {

            string connection = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
            using (SqlConnection con = new SqlConnection(connection))
            {
                SqlCommand command = new SqlCommand("StudentList", con);
                command.CommandType = System.Data.CommandType.StoredProcedure;
                SqlDataAdapter dataAdapter = new SqlDataAdapter();
                dataAdapter.SelectCommand = command;
                con.Open();
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);
                GridView1.DataSource = dataTable;
                GridView1.DataBind();
                con.Close();


            }
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {

            string connection = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
            using (SqlConnection con = new SqlConnection(connection))
            {
                using (SqlCommand command = new SqlCommand("Select * from Student"))
                {
                    using (SqlDataAdapter dataAdapter=new SqlDataAdapter())
                    {
                        command.Connection = con;
                        dataAdapter.SelectCommand = command;
                        using(DataTable dataTable=new DataTable())
                        {
                            dataAdapter.Fill(dataTable);
                            using (XLWorkbook wb= new XLWorkbook())
                            {
                                wb.Worksheets.Add(dataTable, "Student");
                                Response.Clear();
                                Response.Buffer = true;
                                Response.Charset = "";
                                Response.ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                Response.AddHeader("content-disposition","attachment;filename=Student.xlsx");
                                using(MemoryStream ms=new MemoryStream())
                                {
                                    wb.SaveAs(ms);
                                    ms.WriteTo(Response.OutputStream);
                                    Response.Flush();
                                    Response.End();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}