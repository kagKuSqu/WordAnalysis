using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordAnalysis;

namespace UnitTestProject4WordAnalysis
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var mathForm = Form1.ParseMathForm("1+1+1");
            Assert.AreEqual(mathForm,"3");
        }
    }
}
