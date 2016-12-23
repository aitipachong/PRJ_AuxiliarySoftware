using PRJ_DataImport.BLL;
using PRJ_DataImport.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PRJ_DataImport
{
    public partial class FrmMain : Form
    {
        private bool isSelectedExcel = false;       //是否选择了“Excel”文件
        private string selectedFilePath = string.Empty;

        private string yearmonth = string.Empty;
        private string sheetname = string.Empty;

        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            this.InitFrmControls();
        }

        private void InitFrmControls()
        {
            this.textBox1.Enabled = isSelectedExcel;
            this.comboBox1.Items.Clear();
            this.comboBox1.Enabled = isSelectedExcel;
            this.button1.Enabled = isSelectedExcel;
            this.button2.Enabled = isSelectedExcel;
            this.toolStripProgressBar1.Visible = isSelectedExcel;
        }

        private void 打开文件OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "请选择文件";
            openFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx|文本文件(*.txt)|*.txt";
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = openFileDialog.FileName;
                int filterIndex = openFileDialog.FilterIndex;
                switch(filterIndex)
                {
                    case 1:
                        isSelectedExcel = true;
                        this.InitFrmControls();     //初始化控件
                        //1.根据当前年月生成textbox1控件内容
                        this.textBox1.Text = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00");
                        //2.获取Excel文件的所有Sheet页名称,并加入ComboBox控件中
                        List<string> sheetNames = ExcelHelper.GetExcelTablesName(selectedFilePath, ExcelHelper.ExcelType.Excel2007);
                        if(sheetNames != null && sheetNames.Count > 0)
                        {
                            this.comboBox1.Items.Clear();
                            for (int i = 0; i < sheetNames.Count; i++)
                            {
                                if (sheetNames[i].IndexOf("_xlnm#_FilterDatabase") <= 0)
                                {
                                    this.comboBox1.Items.Add(sheetNames[i]);
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Program.ImportStop = true;
        }

        private void 退出XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.ImportStop = true;
            Thread.Sleep(1000);
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.textBox1.Text))
            {
                MessageBox.Show("年月不能为空!");
                return;
            }
            if (string.IsNullOrEmpty(this.comboBox1.Text))
            {
                MessageBox.Show("数据源不能为空！");
                return;
            }

            this.yearmonth = this.textBox1.Text;
            this.sheetname = this.comboBox1.Text;

            Thread th = new Thread(new ThreadStart(StartImportKeyroles));
            th.Start();
            this.toolStripStatusLabel1.Text = "开始导入“关键角色”数据.";
        }

        private void StartImportKeyroles()
        {
            try
            {
                ImportKeyrolesData importLogic = new ImportKeyrolesData(this, this.selectedFilePath, yearmonth, sheetname);
                importLogic.StartImportKeyrolesData();
            }
            catch(Exception ex)
            {
                MessageBox.Show("出现异常，异常为：\r\n" + ex.Message + "\r\n\r\n导入终止!!!", "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}