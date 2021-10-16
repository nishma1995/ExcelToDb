using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelToDb
{
    public partial class UploadFile : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            string name;
            string email;
            string phone;
            string gender;
            string path = Path.GetFileName(FileUpload1.FileName);
            path = path.Replace(" "," ");
            FileUpload1.SaveAs(Server.MapPath("~/Excel File/")+path);
            string ExcelPath = Server.MapPath("~/Excel File/") + path;
            OleDbConnection dbConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+ ExcelPath + "; Extended Properties=Excel 12.0 ;");
            dbConnection.Open();
            OleDbCommand command = new OleDbCommand("select  * from [Sheet1$]",dbConnection);
            OleDbDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                name = dataReader[0].ToString();
                email= dataReader[1].ToString();
                phone=dataReader[2].ToString();
                gender= dataReader[3].ToString();
                saveData(name, email, phone, gender);

            }
        }
        private void saveData(string name1, string email1, string phone1, string gender1)
        {
            string connection = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
            using (SqlConnection con = new SqlConnection(connection))
            {
                con.Open();
                SqlCommand command = new SqlCommand("StudentSave", con);
                command.CommandType = System.Data.CommandType.StoredProcedure;
                command.Parameters.AddWithValue("Name", name1);
                command.Parameters.AddWithValue("Email", email1);
                command.Parameters.AddWithValue("Phone", phone1);
                command.Parameters.AddWithValue("Gender", gender1);
                command.ExecuteNonQuery();
                con.Close();

            }
        }

        protected void btnView_Click(object sender, EventArgs e)
        {
            Response.Redirect("View.aspx");
        }
       
        
    }
}