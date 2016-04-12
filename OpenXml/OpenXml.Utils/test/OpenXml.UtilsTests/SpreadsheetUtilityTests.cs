using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skygp.OpenXml;
using Skygp.OpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Skygp.OpenXml.Tests
{
    [TestClass()]
    public class SpreadsheetUtilityTests
    {
        [TestMethod()]
        public void SpreadsheetUtilityTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ExportTest()
        {
            SpreadsheetUtility<Test> su = new SpreadsheetUtility<Test>(Test.GetDatas());
            su.Export("cc.xlsx");
        }

        [TestMethod()]
        public void ExportTestUseTemplate()
        {
            try
            {
                SpreadsheetUtility<Test> su = new SpreadsheetUtility<Test>(Test.GetDatas());
                su.TemplateFile = @"D:\GithubRep\OpenXml\OpenXml.Utils\test\OpenXml.UtilsTests\TestFile\PlanSapDataT.xlsx";
                su.TemplateRowIndexContent = 1;
                su.GenerateTableHead = false;
                su.TemplateUseLastRowStyle = true;
                su.Export("tt.xlsx");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }
    }

    /// <summary>
    /// 测试类
    /// </summary>
    public class Test
    {
        public static List<Test> GetDatas()
        {
            List<Test> list = new List<Test>() {
                new Test { Id = 1, Name = "理雷", Value = 111, Birthday = DateTime.Now, Sex = true, Score= 11m },
                new Test{ Id = 2, Name = "寒梅", Value = 122, Content="测试内容11222按风格为测试内容11222按风格为测试内容11222按风格为" }
            };
            return list;
        }

        /// <summary>
        /// Id
        /// </summary>
        [NotOpenXml]
        public int Id { get; set; }
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Value
        /// </summary>
        public double Value { get; set; }

        /// <summary>
        /// Birthday
        /// </summary>
        [DisplayOpenXml("生日", "yyyy-MM-dd")]
        [ColumnOpenXml(CustomWidth = 30)]
        public DateTime Birthday { get; set; }
        /// <summary>
        /// Sex
        /// </summary>
        [DisplayOpenXml("性别", BoolTrueText = "男", BoolFalseText = "女")]
        [ColumnOpenXml(CustomWidth = 30)]
        public bool Sex { get; set; }
        /// <summary>
        /// Score
        /// </summary>
        [ColumnOpenXml(CustomWidth = 30)]
        public decimal? Score { get; set; }

        /// <summary>
        /// Content
        /// </summary>
        [ColumnOpenXml(CustomWidth = 60)]
        public string Content { get; set; }
    }
}