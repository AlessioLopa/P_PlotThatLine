using Microsoft.VisualStudio.TestTools.UnitTesting;
using P_PlotThatLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace P_PlotThatLine.Tests
{
    [TestClass()]
    public class Form1Tests
    {
        [TestMethod()]
        public void importCryptoFromCSVTest()
        {
            // Arrange
            var form = new Form1();
            string path = "C:\\Users\\pu41ecx\\Documents\\Github\\P_PlotThatLine\\Data\\bitcoin_2010-01-01_2024-08-28.xlsx";

            // Act
            List<Cryptocurrency> data = form.importCryptoFromCSV(path, "BTC");

            // Assert
            Assert.IsNotNull(data);
            Assert.AreEqual("BTC", data.First().name);
            Assert.IsTrue(data.Count() > 0);
        }
    }
}