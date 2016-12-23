using PRJ_DataImport.Common.DBUtility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PRJ_DataImport.BLL
{
    /// <summary>
    /// 导入“关键角色”业务处理类
    /// </summary>
    public class ImportKeyrolesData
    {
        public FrmMain FrmMain { get; set; }
        public string ExcelPath { get; set; }
        public string YearMonth { get; set; }
        public string SheetName { get; set; }

        private DataTable op_item_t = null;

        public ImportKeyrolesData(FrmMain from, string excelPath, string yearmonth, string sheetname)
        {
            this.FrmMain = from;
            this.ExcelPath = excelPath;
            this.YearMonth = yearmonth;
            this.SheetName = sheetname;
        }

        #region 委托回调FrmMain控件

        public delegate void RefreshPrompt(string message);
        public void ShowPrompt(string message)
        {
            this.FrmMain.toolStripStatusLabel1.Text = message;
        }



        #endregion

        #region 调用委托的方法
        /// <summary>
        /// 刷新“消息”
        /// </summary>
        /// <param name="message"></param>
        public void RefreshMessage(string message)
        {
            RefreshPrompt refreshMsg = new RefreshPrompt(ShowPrompt);
            this.FrmMain.Invoke(refreshMsg, message);
        }

        #endregion

        /// <summary>
        /// 开始导入“关键角色”数据
        /// </summary>
        public void StartImportKeyrolesData()
        {
            try
            {
                //1.加载“op_item_t”数据表数据DataTable对象；
                RefreshMessage("开始获取“op_item_t”数据...");
                this.OP_ITEM_T_DATA();
                RefreshMessage("获取“op_item_t”数据完成.");



            }
            catch(Exception ex)
            {
                RefreshMessage("");
                throw ex;
            }
        }

        /// <summary>
        /// 加载数据库中的op_item_t表数据到DataTable对象
        /// </summary>
        /// <returns></returns>
        private bool OP_ITEM_T_DATA()
        {
            bool result = false;
            try
            {
                string sql = "select * from op_item_t";
                MySqlHelper mysql = new MySqlHelper(Program.MySqlConnectionString);
                DataSet ds = mysql.GetDataSet(CommandType.Text, sql, null);
                if(ds != null && ds.Tables != null && ds.Tables.Count > 0)
                {
                    this.op_item_t = ds.Tables[0];
                    result = true;
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
