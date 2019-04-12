using System;
using TfsSyncHelperForVba;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TfsSyncHelperForVbaTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var syncHelper = new SyncHelper();
            syncHelper.Add(StoreType.Excel, 1, "Title 1e", "InternalId 1e", DateTime.Now);
            syncHelper.Add(StoreType.Excel, 2, "Title 2e", "InternalId 2e", DateTime.Now);
            syncHelper.Add(StoreType.Excel, 3, "Title 3e", "InternalId 3e", DateTime.Now);

            syncHelper.Add(StoreType.Outlook, 1, "Title 1o", "InternalId 1o", new DateTime(2005, 1, 1));
            syncHelper.Add(StoreType.Outlook, 2, "Title 2o", "InternalId 2o", new DateTime(2025, 1, 1));
            syncHelper.Add(StoreType.Outlook, 4, "Title 4o", "InternalId 4o", new DateTime(2025, 1, 1));

            var result = syncHelper.GenerateSyncDictionary();

            Assert.AreEqual(4, result.Count);

            Assert.AreEqual(OperationType.ToOutlook, result[1].OperationType);
            Assert.AreEqual("Title 1e", result[1].TfsActivity.Title);

            Assert.AreEqual(OperationType.ToExcel, result[2].OperationType);
            Assert.AreEqual("Title 2o", result[2].TfsActivity.Title);

            Assert.AreEqual(OperationType.NewToOutlook, result[3].OperationType);
            Assert.AreEqual("Title 3e", result[3].TfsActivity.Title);

            Assert.AreEqual(OperationType.NewToExcel, result[4].OperationType);
            Assert.AreEqual("Title 4o", result[4].TfsActivity.Title);

        }
    }
}
