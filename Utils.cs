namespace LostTech.Stack.WindowManagement {
    using System;
    using System.Diagnostics;
    static class Utils {
        public static void ReportAsWarning(this Exception exception)
            => Debug.WriteLine("WARN: " + exception);
    }
}
