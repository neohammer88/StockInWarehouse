/* Title : BMS Stock program
 * Version : 1.2
 * Language : C#
 * Programmer : Tom Rho
 * Date : 23/07/2018
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace $safeprojectname$
{
    public partial class Main : Form
    {

        public Main()
        {
            InitializeComponent();
        }

        public static string user;

        //Get Data from the database and Make DataSet
        private DataSet GetData()
        {
            string connStr = "Data Source = (local); Initial Catalog = master; Integrated Security = true";
            SqlConnection conn = new SqlConnection(connStr);
            conn.Open();

            DataSet ds = new DataSet();
            string sql = "SELECT * FROM Stock";
            SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
            adapter.Fill(ds);
            conn.Close();

            return ds;
        }

        //Exit the program
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Login login = new Login();
            login.ShowDialog();
            lblUser.Text = Login.user1;
        }

        //Call Form2 to Search stocks
        private void btnSearch_Click(object sender, EventArgs e)
        {
                user = lblUser.Text;
                Search f2 = new Search();
                f2.Show();
        }

        //Call Form3 to Save stocks
        private void btnStock_Click(object sender, EventArgs e)
        {
            //if (string.IsNullOrEmpty(txtBoxUser.Text))
            //{
            //    MessageBox.Show("Please Input User Name", "ERROR", MessageBoxButtons.OK);
            //}
            //else
            //{
                user = lblUser.Text;
                Save f3 = new Save();
                f3.Show();
            //}
        }

        //Call Form4 to delete stocks
        private void btnDelete1_Click(object sender, EventArgs e)
        {
                user = lblUser.Text;
                Delete f4 = new Delete();
                f4.Show();
        }

        //Call Form6 to move stocks
        private void btnMove_Click(object sender, EventArgs e)
        {
                user = lblUser.Text;
                Move f6 = new Move();
                f6.Show();
        }

        //Extract all data to Excel file
        private void btnExcel_Click(object sender, EventArgs e)
        {
            // Call Data
            DataSet ds = GetData();
            Excel.Application ap = new Excel.Application();
            Excel.Workbook excelWorkBook = ap.Workbooks.Add();

            //Put Data to Excel
            foreach (DataTable dt in ds.Tables)
            {
                Excel.Worksheet ws = excelWorkBook.Sheets.Add();
                ws.Name = dt.TableName;

                //Input Header Name in Excel
                for(int columnHeaderIndex = 1; columnHeaderIndex <= dt.Columns.Count; columnHeaderIndex++)
                {
                    ws.Cells[1, columnHeaderIndex] = dt.Columns[columnHeaderIndex - 1].ColumnName;
                    ws.Cells[1, columnHeaderIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue);
                }

                //Input every data of DataSet
                for(int rowIndex = 0; rowIndex <= (dt.Rows.Count - 1); rowIndex++)
                {
                    for(int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                    {
                        ws.Cells[rowIndex + 2, columnIndex + 1] = dt.Rows[rowIndex].ItemArray[columnIndex].ToString();
                    }
                }
                //Adjust the column width automatically
                ws.Columns.AutoFit();
            }

            //Make a file with extension and saving directory
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            saveFile.Title = "File Saving";
            saveFile.DefaultExt = "xlsx";
            saveFile.Filter = "Xlsx files(*.xlsx)|*.xlsx|Xls files(*.xls)|*.xls";
            saveFile.ShowDialog();

            if(saveFile.FileNames.Length > 0)
            {
                foreach(string filename in saveFile.FileNames)
                {
                    string savePath = filename;
                    if (Path.GetExtension(savePath) == ".xls")
                    {
                        excelWorkBook.SaveAs(savePath, Excel.XlFileFormat.xlWorkbookNormal);
                    }
                    else if (Path.GetExtension(savePath) == ".xlsx")
                    {
                        excelWorkBook.SaveAs(savePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    excelWorkBook.Close(true);
                    ap.Quit();
                }
            }
        }

        //Call History form to search stock's history
        private void btnHistory_Click(object sender, EventArgs e)
        {
                user = lblUser.Text;
                History history = new History();
                history.Show();
        }

        private void btnBulk_Click(object sender, EventArgs e)
        {
            user = lblUser.Text;
            BulkInsert f9 = new BulkInsert();
            f9.Show();
        }
    }
}
