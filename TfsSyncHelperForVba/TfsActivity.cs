using System;
using System.Runtime.InteropServices;

namespace TfsSyncHelperForVba
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TfsActivity
    {
        public string Title { get; set; }
        public string Id { get; set; }
        public DateTime LastEdit { get; set; }
    }
}
