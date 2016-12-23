using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;
using PRJ_DataImport.Common.DBUtility;
using System.Data;

namespace UT_PRJ_DataImport
{
    [TestClass]
    public class UT_MySqlHelper
    {
        /// <summary>
        /// 测试数据库是否链接
        /// </summary>
        [TestMethod]
        public void UT_DatabaseConnecte_V1()
        {
            try
            {
                string strConnection = ConfigurationManager.ConnectionStrings["MysqlConnectionString"].ConnectionString;
                if(string.IsNullOrEmpty(strConnection))
                {
                    Assert.Fail("App.config文件中没有设置名为“MysqlConnectionString”的数据库链接字符串!");
                    return;
                }

                MySqlHelper mysql = new MySqlHelper(strConnection);
                string sql = "select count(*) from tbl_sys_user";
                int num = Convert.ToInt32(mysql.ExecuteScalar(CommandType.Text, sql, null));

                Assert.AreEqual(724, num);
            }
            catch(Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

    }
}
