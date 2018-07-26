using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Windows;


namespace UpdateCustomersTest
{
    [TestClass]
    public class ViewCustomersTests
    {
        [TestMethod]
        public void TestGetDriverList()
        {
            List<UpdateCustomers.ViewCustomers.Driver> EmptyList1 = new List<UpdateCustomers.ViewCustomers.Driver>();
    
            UpdateCustomers.ViewCustomers view = new UpdateCustomers.ViewCustomers();
            string filepath = @"c:\\badfilename.txt";

            CollectionAssert.AreEqual(EmptyList1, view.GetDriverList(filepath));
        }
    }
}
