using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Globalization;
using DevExpress.Drawing.Internal.Fonts.Interop;
namespace _2224802010871_Phan_Phuoc_Hong_Phuc_LAB6
{
    public partial class Form2 : Form
    {
        private SqlConnection conn;
        string maHS, tenHS, diachi, malop;
        double dtb;
        DateTime ngaysinh;
        public Form2()
        {
            InitializeComponent();
        }
        public void connect()
        {
            //Lấy chuỗi kết nối CSDL
            string strcon = "Data Source=LAPTOP-PICGBI40\\HONGPHUC;Initial Catalog=QLHOCSINH;Integrated Security=True;";
            try
            {
                conn = new SqlConnection(strcon);
                // Mở kết nối
                conn.Open();
                //Ko làm gì thì đóng kết nối lại
                conn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Không kết nối được CSDL", "Thông báo",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            connect();
            //Tạo 1 datatable để lấy dữ liệu từ bảng Lớp
            //qua hàm getDSLop()
            DataTable table = getDSLop();
            //Đổ dữ liệu lên combobox
            cmbLop.DataSource = table;
            //Nội dung hiển thị lên combobox
            cmbLop.DisplayMember = "TenLop";
            //giá trị truy xuất combobox
            cmbLop.ValueMember = "MaLop";
            dgvHocSinh.DataSource = getDSHocSinh();
            DinhDangLuoi();
        }
        public DataTable getDSLop()
        {
            string strSQL = "Select * from Lop";
            SqlDataAdapter adapter;
            adapter = new SqlDataAdapter(strSQL, conn);
            DataSet dataset = new DataSet();
            try
            {
                adapter.Fill(dataset);
                return dataset.Tables[0];
            }
            catch
            {
                return null;
            }
        }
        private DataTable getDSHocSinh()
        {
            string str = "Select h.MaHS, h.TenHS, h.NgaySinh, h.DiaChi, h.DTB,l.TenLop From HOCSINH h, LOP l Where h.MaLop = l.MaLop";
            SqlDataAdapter adapter = new SqlDataAdapter(str, conn);
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet);
            return dataSet.Tables[0];
        }
        private void DinhDangLuoi()
        {
            dgvHocSinh.ReadOnly = true;
            dgvHocSinh.Columns[0].HeaderText = "Mã HS";
            dgvHocSinh.Columns[0].Width = 70;
            dgvHocSinh.Columns[1].HeaderText = "Tên HS";
            dgvHocSinh.Columns[1].Width = 100;
            dgvHocSinh.Columns[2].HeaderText = "Ngày sinh";
            dgvHocSinh.Columns[2].Width = 90;
            dgvHocSinh.Columns[3].HeaderText = "Địa chỉ";
            dgvHocSinh.Columns[3].Width = 150;
            dgvHocSinh.Columns[4].HeaderText = "Điểm TB";
            dgvHocSinh.Columns[4].Width = 80;
            dgvHocSinh.Columns[5].HeaderText = "Lớp";
            dgvHocSinh.Columns[5].Width = 80;
        }

        public bool Check()
        {
            string strSQL = "SELECT COUNT(*) FROM HocSinh";
            try
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                int count = (int)cmd.ExecuteScalar(); // Sử dụng ExecuteScalar cho COUNT
                return count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối: " + ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }
        private void dgvHocSinh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewSelectedCellCollection cell = dgvHocSinh.SelectedCells;
            if (cell.Count > 0)
            {
                DataGridViewRow row = dgvHocSinh.Rows[e.RowIndex];
                txtMaHS.Text = row.Cells["MaHS"].Value.ToString();
                txtTenHS.Text = row.Cells["TenHS"].Value.ToString();
                if (row.Cells["NgaySinh"].Value.ToString().Length > 0)
                    dtpNgaySinh.Text = row.Cells["NgaySinh"].Value.ToString();
                txtDiaChi.Text = row.Cells["DiaChi"].Value.ToString();
                cmbLop.Text = row.Cells["TenLop"].Value.ToString();
                txtDiemTB.Text = row.Cells["DTB"].Value.ToString();
            }
        }
        private void ClearStudentData()
        {
            txtTenHS.Text = "";
            dtpNgaySinh.Value = DateTime.Now;
            txtDiaChi.Text = "";
            txtDiemTB.Text = "";
            btnThem.Text = "Lưu";
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
            maHS = txtMaHS.Text;
            tenHS = txtTenHS.Text;
            ngaysinh = dtpNgaySinh.Value; //DateTimePicker
            diachi = txtDiaChi.Text;
            malop = cmbLop.SelectedValue.ToString();
            dtb = Convert.ToDouble(txtDiemTB.Text);
            if (txtMaHS.Text == "")
            {
                errorProvider1.SetError(txtMaHS, "Hãy nhập mã hs!");
                return;
            }
            if (txtTenHS.Text == "")
            {
                errorProvider1.SetError(txtTenHS, "Hãy nhập tên hs!");
                return;
            }
            if (dtb < 0 || dtb > 10)
            {
                errorProvider1.SetError(txtDiemTB, "DTB > 0 va < 10");
                return;
            }
            if (DateTime.Now.Year - ngaysinh.Year < 15 || DateTime.Now.Year - ngaysinh.Year > 20)
            {
                errorProvider1.SetError(dtpNgaySinh, "15 -> 20");
                return;
            }
            try
            {
                conn.Open();
                string str = "insert into HocSinh VALUES('" +
                maHS + "',N'" + tenHS + "','" +
                ngaysinh + "',N'" + diachi + "'," +
                dtb + ",'" + malop + "')";
                SqlCommand cmd1 = new SqlCommand(str, conn);
                cmd1.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Thêm dữ liệu thành công");
                txtMaHS.Text = "";
                txtTenHS.Text = "";
                dtpNgaySinh.Text = "";
                txtDiaChi.Text = "";
                txtDiemTB.Text = "";
                cmbLop.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm dữ liệu: " + ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            dgvHocSinh.DataSource = getDSHocSinh();
        }

        private void btnXoa_Click_1(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                string strSQL = "delete from HocSinh where maHS ='" + maHS + "'";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Xóa dữ liệu thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xóa bị lỗi" + ex.ToString());
            }
            dgvHocSinh.DataSource = getDSHocSinh();
        }

        private void BtnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(maHS))
            {
                MessageBox.Show("Vui lòng tìm kiếm học sinh trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string updatedTenHS = txtTenHS.Text;
            DateTime updatedNgaySinh = dtpNgaySinh.Value;
            string updatedDiaChi = txtDiaChi.Text;
            double updatedDTB = double.Parse(txtDiemTB.Text, CultureInfo.InvariantCulture);
            string updatedMaLop = cmbLop.SelectedValue.ToString();
            try
            {
                conn.Open();
                string str = "UPDATE HocSinh SET TenHS = @tenHS, NgaySinh = @ngaysinh, DiaChi = @diachi, DTB = @dtb, MaLop = @malop WHERE MaHS = @maHS";

                SqlCommand cmd = new SqlCommand(str, conn);
                cmd.Parameters.AddWithValue("@maHS", maHS);
                cmd.Parameters.AddWithValue("@tenHS", updatedTenHS);
                cmd.Parameters.AddWithValue("@ngaysinh", updatedNgaySinh);
                cmd.Parameters.AddWithValue("@diachi", updatedDiaChi);
                cmd.Parameters.AddWithValue("@dtb", updatedDTB);
                cmd.Parameters.AddWithValue("@malop", updatedMaLop);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Cập nhật dữ liệu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dgvHocSinh.DataSource = getDSHocSinh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi cập nhật dữ liệu: " + ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void txtMaHS_TextChanged_1(object sender, EventArgs e)
        {
            maHS = txtMaHS.Text;
            if (string.IsNullOrEmpty(maHS))
            {
                ClearStudentData();
                return;
            }
            try
            {
                conn.Open();
                string strSQL = "SELECT * FROM HocSinh WHERE maHS = '" + maHS + "'";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.Parameters.AddWithValue("@maHS", maHS);
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.HasRows)
                {
                    reader.Read();

                    tenHS = !reader.IsDBNull(1) ? reader.GetString(1) : string.Empty;
                    ngaysinh = !reader.IsDBNull(2) ? reader.GetDateTime(2) : DateTime.Now;
                    diachi = !reader.IsDBNull(3) ? reader.GetString(3) : string.Empty;
                    dtb = !reader.IsDBNull(4) ? reader.GetDouble(4) : 0.0;
                    malop = !reader.IsDBNull(5) ? reader.GetString(5) : string.Empty;

                    txtTenHS.Text = tenHS;
                    dtpNgaySinh.Value = ngaysinh;
                    txtDiaChi.Text = diachi;
                    txtDiemTB.Text = dtb.ToString(CultureInfo.InvariantCulture);
                    cmbLop.SelectedValue = malop;
                }
                else
                {
                    ClearStudentData();
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lấy dữ liệu học sinh: " + ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Import Excel";
            openFileDialog.Filter = "Excel (*.xlsx) | *.xlsx|Excel 2003 (*.xls)|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ImportExcel(openFileDialog.FileName);
                    MessageBox.Show("Nhập file thành công!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nhập file không thành công!\n" + ex.Message);
                }
            }
        }

        private void ExportExcel(string path)
        {
            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for(int i = 0; i < dgvHocSinh.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dgvHocSinh.Columns[i].HeaderText;
            }
            for(int i = 0; i < dgvHocSinh.Rows.Count; i++)
            {
                for(int j = 0; j < dgvHocSinh.Columns.Count; j++)
                {
                    application.Cells[i + 2, j + 1] = dgvHocSinh.Rows[i].Cells[j].Value;
                }
            }
            application.Columns.AutoFit();
            application.ActiveWorkbook.SaveCopyAs(path);
            application.ActiveWorkbook.Saved = true;

        }
        private void ImportExcel(string path)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet  = package.Workbook.Worksheets[0];
                DataTable data = new DataTable();
                for(int i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column; i++)
                {
                    data.Columns.Add(excelWorksheet.Cells[1,i].Value.ToString());
                }
                for(int i = excelWorksheet.Dimension.Start.Row + 1; i <= excelWorksheet.Dimension.End.Row; i++)
                {
                    List<string> listRows = new List<string>();
                    for(int j = excelWorksheet.Dimension.Start.Column; j <= excelWorksheet.Dimension.End.Column; j++)
                    {
                        listRows.Add(excelWorksheet.Cells[i, j].Value.ToString());
                    }
                    data.Rows.Add(listRows.ToArray());
                }
                dgvHocSinh.DataSource = data;
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx) | *.xlsx|Excel 2003 (*.xls)|*.xls";
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("Xuất file thành công!");   
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Xuất file không thành công!\n" + ex.Message);
                }            
            }
        }
    }
}
