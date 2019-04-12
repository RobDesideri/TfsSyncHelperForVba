using System.Runtime.InteropServices;

namespace TfsSyncHelperForVba
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class SyncOperation
    {
        public TfsActivity TfsActivity { get; set; }

        public OperationType OperationType { get; set; }
    }
}
