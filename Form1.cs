using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelIT_Tracker
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void dataclear_btn_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Veriler Silinsin mi?", "Veri Silme", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                MessageBox.Show("Tüm Veriler Silindi!");
                if (datagrid_IT.DataSource is DataTable dt)
                {
                    dt.Clear();
                }
                else
                {
                    datagrid_IT.DataSource = null;
                    datagrid_IT.Rows.Clear();
                }

            }
            else if (dialogResult == DialogResult.No)
            {
             
            }
        }
        private void trasnfer_btn_Click(object sender, EventArgs e)
        {
            if (envanter_text.Text == "" || serino_text.Text == "" || pcad_text.Text == "" || marka_text.Text == "" || model_text.Text == "" || islemci_text.Text == "" ||
            nesil_combo.SelectedItem == null || ramk_combo.SelectedItem == null || ddr_combo.SelectedItem == null || disktur_combo.SelectedItem == null ||
            disk_combo.SelectedItem == null || isletim_combo.SelectedItem == null || surum_combo.SelectedItem == null || kulanici_text.Text == "" || lokasyon_text.Text == "" || birim_text.Text == null || diger_rich.Text == null)
            {
                MessageBox.Show("Lütfen Boş Alan Bırakmayınız!");
            }
            else
            {
                datagrid_IT.Rows.Add(envanter_text.Text, serino_text.Text, pcad_text.Text, marka_text.Text, model_text.Text, islemci_text.Text,
                nesil_combo.SelectedItem.ToString(), ramk_combo.SelectedItem.ToString(), ddr_combo.SelectedItem.ToString(), disktur_combo.SelectedItem.ToString(),
                disk_combo.SelectedItem.ToString(), isletim_combo.SelectedItem.ToString(),
                surum_combo.SelectedItem.ToString(), kulanici_text.Text, lokasyon_text.Text, birim_text.Text, diger_rich.Text);
                temizle();
            }
        }
        private void clr_btn_Click(object sender, EventArgs e)
        {
            temizle();
        }
        private void exprtexc_btn_Click(object sender, EventArgs e)
        {
           Export(datagrid_IT);
        }
        private void Main_Load(object sender, EventArgs e)
        {

            //Nesil
            nesil_combo.Items.Add("1.Nesil");
            nesil_combo.Items.Add("2.Nesil");
            nesil_combo.Items.Add("3.Nesil");
            nesil_combo.Items.Add("4.Nesil");
            nesil_combo.Items.Add("5.Nesil");
            nesil_combo.Items.Add("6.Nesil");
            nesil_combo.Items.Add("7.Nesil");
            nesil_combo.Items.Add("8.Nesil");
            nesil_combo.Items.Add("9.Nesil");
            nesil_combo.Items.Add("10.Nesil");
            nesil_combo.Items.Add("11.Nesil");
            nesil_combo.Items.Add("12.Nesil");
            nesil_combo.Items.Add("13.Nesil");
            nesil_combo.Items.Add("14.Nesil");

            //Ram Kapasite
            ramk_combo.Items.Add("2 GB");
            ramk_combo.Items.Add("4 GB");
            ramk_combo.Items.Add("8 GB");
            ramk_combo.Items.Add("16 GB");
            ramk_combo.Items.Add("32 GB");
            ramk_combo.Items.Add("64 GB");

            //Ram DDR
            ddr_combo.Items.Add("DDR1");
            ddr_combo.Items.Add("DDR2");
            ddr_combo.Items.Add("DDR3");
            ddr_combo.Items.Add("DDR4");
            ddr_combo.Items.Add("DDR5");

            //Disk Türü
            disktur_combo.Items.Add("HDD");
            disktur_combo.Items.Add("SSD");
            disktur_combo.Items.Add("M.2 SSD");

            //Disk 
            disk_combo.Items.Add("120GB");
            disk_combo.Items.Add("256GB");
            disk_combo.Items.Add("500GB");
            disk_combo.Items.Add("1TB");
            disk_combo.Items.Add("2TB");
            disk_combo.Items.Add("4TB");
            disk_combo.Items.Add("8TB");
            disk_combo.Items.Add("16TB");

            //İşletim Sistemi
            isletim_combo.Items.Add("Diğer");
            isletim_combo.Items.Add("Windows7");
            isletim_combo.Items.Add("Windows8");
            isletim_combo.Items.Add("Windows10");
            isletim_combo.Items.Add("Windows11");

            //Sürüm
            surum_combo.Items.Add("Home");
            surum_combo.Items.Add("Pro");
            surum_combo.Items.Add("Education");
            surum_combo.Items.Add("Enterprise ");
            surum_combo.Items.Add("Pro For Workstations");
            surum_combo.Items.Add("Server");

            
        }
        void temizle()
        {
            envanter_text.Clear();
            serino_text.Clear();
            pcad_text.Clear();
            marka_text.Clear();
            model_text.Clear();
            islemci_text.Clear();
            nesil_combo.SelectedItem = null;
            ramk_combo.SelectedItem = null;
            ddr_combo.SelectedItem = null;
            disktur_combo.SelectedItem = null;
            disk_combo.SelectedItem = null; 
            isletim_combo.SelectedItem=null;
            surum_combo.SelectedItem=null;
            kulanici_text.Clear();
            lokasyon_text.Clear();
            birim_text.Clear();
            diger_rich.Clear();

        }
        private void Export(DataGridView dgv)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet 1");

                    for (int i = 0; i < dgv.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = dgv.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Rows[i].Cells[j].Value != null)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = dgv.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Veriler Excel'e aktarıldı");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
        private void search_textBox_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                string searchValue = search_textBox.Text.Trim().ToLower();

                foreach (DataGridViewRow row in datagrid_IT.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool found = false;

                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null && cell.Value != DBNull.Value) // NULL kontrolü
                        {
                            string cellValue = "";

                            try
                            {
                                if (cell.Value is decimal || cell.Value is double || cell.Value is float || cell.Value is int || cell.Value is long)
                                {
                                    cellValue = Convert.ToDouble(cell.Value).ToString(); 
                                }
                                else
                                {
                                    cellValue = cell.Value.ToString().ToLower(); 
                                }
                            }
                            catch
                            {
                                cellValue = "";
                            }

                            if (cellValue.Contains(searchValue))
                            {
                                found = true;
                                break;
                            }
                        }
                    }

                    row.Visible = found;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen Doğru Değer Giriniz!");
            }

        }
    }
}
