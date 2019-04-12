using System.Runtime.InteropServices;

namespace TfsSyncHelperForVba
{

    [ComVisible(true)]
    public enum StoreType
    {
        Excel = 1,
        Outlook = 2
    }

    [ComVisible(true)]
    public enum OperationType
    {
        NewToExcel = 1,
        NewToOutlook = 2,
        ToExcel = 3,
        ToOutlook = 4
    }
}