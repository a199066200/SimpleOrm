using NUnit.Framework;
using SimpleOrm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    [TestFixture]
    public class TestClass
    {
        private string path = @"C:\Users\Administrator\Desktop\TestExcel.xlsx";
        [Test]
        public void TestWriteRead()
        {
            var list = new List<Test2>();
            list.Add(new Test2 { Count = 1, Log = "Test1"});
            list.Add(new Test2 { Count = 2, Log = "周扬测试"});
            var table = SimpleOrm.SimpleOrm.BuildDataTable(list);
            ExcelOperator.WriteOnNewSheet(table, path);
            var table2 = ExcelOperator.QueryAll(path, typeof(Test2).Name);
            var result = SimpleOrm.SimpleOrm.Read(table2);
            Assert.AreEqual(2, result.Count());
            Assert.AreEqual("1", result.First().Count);
            Assert.AreEqual("周扬测试", result.Last().Log);
        }
    }

    public class Test2 {
        public int Count { get; set; }
        public string Log { get; set; }
    }
}
