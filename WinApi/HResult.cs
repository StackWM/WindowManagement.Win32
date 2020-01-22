// ReSharper disable InconsistentNaming
namespace LostTech.Stack.WindowManagement.WinApi {
    public enum HResult:uint {
        TYPE_E_ELEMENTNOTFOUND = 0x8002802B,
        ERROR_INVALID_CURSOR_HANDLE = 0x8007057A,
    }

    public static class HResultExtensions {
        public static bool EqualsCode(this HResult result, int code) => result == (HResult)code;
    }
}