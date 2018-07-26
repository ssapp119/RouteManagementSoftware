using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UpdateCustomersTest
{
    [TestClass]
    class UpdateRoutesTests
    {


        [TestMethod]
        public void TestGetRouteList()
        {
            
                List<UpdateCustomers.ViewCustomers.Routes> EmptyList1 = new List<UpdateCustomers.ViewCustomers.Routes>();

                UpdateCustomers.UpdateRoutes routes = new UpdateCustomers.UpdateRoutes();
                string filepath = @"c:\\badfilename.txt";

                CollectionAssert.AreEqual(EmptyList1, routes.GetRouteList(filepath));
            
        }


    }
}
