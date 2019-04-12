using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace TfsSyncHelperForVba
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class SyncHelper
    {
        private Dictionary<int, TfsActivity> excelData = new Dictionary<int, TfsActivity>();
        private Dictionary<int, TfsActivity> outlookData = new Dictionary<int, TfsActivity>();

        public void Add(StoreType storeType, int wiId, string title, string id, DateTime lastEdit)
        {
            switch (storeType)
            {
                case StoreType.Excel:
                    excelData.Add(wiId, new TfsActivity { Id = id, LastEdit = lastEdit, Title = title });
                    break;
                case StoreType.Outlook:
                    outlookData.Add(wiId, new TfsActivity { Id = id, LastEdit = lastEdit, Title = title });
                    break;
                default:
                    throw new ArgumentException("Bad value on storeType parameter");
            }
        }

        [ComVisible(true)]
        public Dictionary<int, SyncOperation> GenerateSyncDictionary()
        {
            if (excelData.Count == 0 || outlookData.Count == 0)
            {
                throw new InvalidOperationException("No data are added.");
            }

            var excelKeys = this.excelData.Keys.ToList();
            var outlookKeys = this.outlookData.Keys.ToList();

            var idsOnlyOnExcel = excelKeys.Except(outlookKeys);
            var idsOnlyOnOutlook = outlookKeys.Except(excelKeys);
            var sharedIds = excelKeys.Intersect(outlookKeys);
            List<int> allIds = excelKeys.Union(outlookKeys).ToList();
            allIds.Sort();

            Dictionary<int, SyncOperation> syncDictionary = new Dictionary<int, SyncOperation>();

            for (int i = 0; i < allIds.Count; i++)
            {
                var thisId = allIds[i];
                OperationType thisOperationType;
                TfsActivity thisTfsActivity;

                if (sharedIds.Contains(thisId))
                {
                    var excelItemLastEdit = excelData[thisId].LastEdit;
                    var outlookItemLastEdit = outlookData[thisId].LastEdit;

                    if (excelItemLastEdit >= outlookItemLastEdit)
                    {
                        thisOperationType = OperationType.ToOutlook;
                        thisTfsActivity = excelData[thisId];
                    }
                    else
                    {
                        thisOperationType = OperationType.ToExcel;
                        thisTfsActivity = outlookData[thisId];
                    }
                }
                else if (idsOnlyOnExcel.Contains(thisId))
                {
                    thisOperationType = OperationType.NewToOutlook;
                    thisTfsActivity = excelData[thisId];
                }
                else if (idsOnlyOnOutlook.Contains(thisId))
                {
                    thisOperationType = OperationType.NewToExcel;
                    thisTfsActivity = outlookData[thisId];
                }
                else
                {
                    throw new InvalidOperationException("Found an ID not identified in both Excel or Outlook data");
                }

                syncDictionary.Add(allIds[i], new SyncOperation { OperationType = thisOperationType, TfsActivity = thisTfsActivity });
            }

            return syncDictionary;
        }
    }
}
