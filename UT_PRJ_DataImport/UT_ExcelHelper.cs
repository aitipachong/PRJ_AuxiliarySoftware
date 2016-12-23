using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PRJ_DataImport.Common;
using System.Collections.Generic;

namespace UT_PRJ_DataImport
{
    [TestClass]
    public class UT_ExcelHelper
    {
        /// <summary>
        /// 获取指定Excel文件的工作表名称
        /// </summary>
        [TestMethod]
        public void UT_GetExcelTablesName_V1()
        {
            string excelPath = @"E:\WORKs\项目管理IT系统\01 Documents\00.4 Bug单\01 导入”关键绝色“到数据库对应表\项目关键角色清单、能力评估及规划（11月）-V2.0.xlsx";
            try
            {
                if(!File.Exists(excelPath))
                {
                    Assert.Fail(string.Format("Excel文件不存在,打开路径为：{0}", excelPath));
                    return;
                }

                List<string> tableNames = ExcelHelper.GetExcelTablesName(excelPath, ExcelHelper.ExcelType.Excel2007);
                Assert.IsTrue(tableNames.Count > 0);
            }
            catch(Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }
    }
}
