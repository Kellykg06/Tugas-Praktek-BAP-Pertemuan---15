using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dapper;
using Excel = Microsoft.Office.Interop.Excel;

namespace Tugas_Praktek_BAP_Pertemuan_15
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dgvData.AutoGenerateColumns = false;
            this.lblBanyakRecordData.Text = "";
            LoadData();
        }
        string connString = @"Data Source=.\sqlexpress; Initial Catalog=DB_IMS; Integrated Security=True;";
        
        private List<IMS> listData = null;
        private List<DataSupplier> listDataSupplier = null;
        private List<IMS> GetListAllDataIMS()
        {
            IEnumerable<IMS> listData = null;
            try
            {
                using (var conn = new SqlConnection(connString))
                {
                   listData = conn.Query<IMS>(@"SELECT K.Nomor, K.Tanggal, K.Supplier, COUNT(L.Quantity) AS Quantity, K.Keterangan FROM TOrder K INNER JOIN OrderDetail L
                                                ON K.Nomor = L.Nomor GROUP BY K.Nomor, K.Tanggal, K.Supplier, K.Keterangan");
                }
            }
            catch (Exception)
            {
                throw;
            }
            return listData?.ToList() ?? null;
        }
        private List<DataSupplier> GetListAllDataDataSupplier()
        {
            IEnumerable<DataSupplier> listDataSupplier = null;
            try
            {
                using (var conn = new SqlConnection(connString))
                {
                    listDataSupplier = conn.Query<DataSupplier>(@"SELECT L.Nomor, L.KodeBarang, L.Quantity, M.NamaBarang, M.Satuan FROM OrderDetail L INNER JOIN Barang M 
                    ON M.KodeBarang = L.KodeBarang ");
                }
            }
            catch (Exception)
            {
                throw;
            }
            return listDataSupplier?.ToList() ?? null;
        }
        private void LoadData()
        {
            try
            {
                listDataSupplier = GetListAllDataDataSupplier();
                listData = GetListAllDataIMS();
                if (listData != null && listData.Any())
                {
                    this.dgvData.DataSource = listData;
                    this.dgvData.Columns[0].DataPropertyName = nameof(IMS.Nomor);
                    this.dgvData.Columns[1].DataPropertyName = nameof(IMS.Tanggal);
                    this.dgvData.Columns[2].DataPropertyName = nameof(IMS.Supplier);
                    this.dgvData.Columns[3].DataPropertyName = nameof(IMS.Quantity);
                    this.dgvData.Columns[4].DataPropertyName = nameof(IMS.Keterangan);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                this.lblBanyakRecordData.Text = $"{ this.dgvData.Rows.Count:n0} Record Data.";
            }
        }

        private void btnExportToTemplate_Click(object sender, EventArgs e)
        {
            if (listData != null && listData.Any())
            {
                try
                {
                    Excel.Application app = new Excel.Application();

                    Excel.Workbook book = app.Workbooks.Add();

                    Excel.Worksheet sheet = book.ActiveSheet as Excel.Worksheet;

                    app.Visible = true;
                    app.WindowState = Excel.XlWindowState.xlMaximized;

                    sheet.Cells[1, 1] = "Data Order";

                    int barisHeader = 3;

                    sheet.Cells[barisHeader, 1] = "Nomor";
                    sheet.Cells[barisHeader, 2] = "Tanggal";
                    sheet.Cells[barisHeader, 3] = "Supplier";
                    sheet.Cells[barisHeader + 1, 2] = "Kode Barang";
                    sheet.Cells[barisHeader + 1, 3] = "Nama Barang";
                    sheet.Cells[barisHeader + 1, 4] = "Quantity";
                    sheet.Cells[barisHeader + 1, 5] = "Satuan";
                    sheet.Range[$"A{barisHeader}", $"C{barisHeader}"].Font.Bold = true;
                    sheet.Range[$"B{barisHeader + 1}", $"E{barisHeader + 1}"].Font.Bold = true;
                    sheet.Range[$"A{barisHeader}", $"C{barisHeader}"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    sheet.Range[$"B{barisHeader + 1}", $"E{barisHeader}"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    for (int i = 0; i < listData.Count; i++)
                    {
                        sheet.Cells[barisHeader + 3, 1] = listData[i].Nomor;
                        sheet.Cells[barisHeader + 3, 2] = listData[i].Tanggal;
                        sheet.Cells[barisHeader + 3, 3] = listData[i].Supplier;
                        sheet.Cells[barisHeader + 3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        barisHeader++;
                        for (int j = 0; j < listDataSupplier.Count; j++)
                        {
                            if (listData[i].Nomor == listDataSupplier[j].Nomor)
                            {
                                sheet.Cells[barisHeader + 3, 2] = listDataSupplier[j].KodeBarang;
                                sheet.Cells[barisHeader + 3, 3] = listDataSupplier[j].NamaBarang;
                                sheet.Cells[barisHeader + 3, 4] = listDataSupplier[j].Quantity;
                                sheet.Cells[barisHeader + 3, 5] = listDataSupplier[j].Satuan;
                                sheet.Cells[barisHeader + 3, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                barisHeader++;
                            }
                        }
                        barisHeader++;
                    }   

                    sheet.Range["A1", "E1"].Font.Bold = true;
                    sheet.Range["A1", "E1"].MergeCells = true;
                    sheet.Range["A1", "E1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    sheet.Columns.AutoFit();
                    sheet.Rows.AutoFit();

                    sheet.Name = "Orders";
                    app.UserControl = true;
                    book.Password = "123456";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Sorry, tidak ada data yang bisa diexport ...", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
