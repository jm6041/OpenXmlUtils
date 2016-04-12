using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Skygp.OpenXml
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertTest();
            
            //Test t1 = new Test { Id = 1, Name = "理雷", Value = 111, Birthday = DateTime.Now, Sex = true, Score = 11m };
            //int hash1 = EqualUtil.PublicInstancePropertiesHashCodeXor(t1);
            //Test t2 = new Test { Id = 1, Name = "理雷", Value = 111, Birthday = DateTime.Now, Sex = true, Score = 11m };
            //bool eq1 = t1.Equals(t2);
            //bool eq = EqualUtil.PublicInstancePropertiesEqual(t1, t2);
        }

        public static void InsertTest()
        {
            List<Test> list = new List<Test>() {
                new Test { Id = 1, Name = "理雷", Value = 111, Birthday = DateTime.Now, Sex = true, Score= 11m },
                new Test{ Id = 2, Name = "寒梅", Value = 122, Content="测试内容11222按风格为测试内容11222按风格为测试内容11222按风格为" }
            };
            SpreadsheetUtility<Test> cc = new SpreadsheetUtility<Test>(list);
            cc.UseDefalutStyle = true;
            cc.Export("cc.xlsx");
        }
    }

    /// <summary>
    /// <see cref="Test"/>
    /// </summary>
    public class Test
    {
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
