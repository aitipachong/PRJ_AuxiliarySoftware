using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PRJ_DataImport
{
    static class Program
    {
        /// <summary>
        /// MySQL数据库链接字符串
        /// </summary>
        public static string MySqlConnectionString { get; set; }

        /// <summary>
        /// 导入停止
        /// </summary>
        public static bool ImportStop { get; set; } = false;

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            /*
              全局变量
                1、MySQL数据库链接字符串
            */
            MySqlConnectionString = ConfigurationManager.ConnectionStrings["MysqlConnectionString"].ConnectionString;

            Application.Run(new FrmMain());
        }
    }
}
