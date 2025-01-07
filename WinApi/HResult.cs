// ReSharper disable InconsistentNaming
namespace LostTech.Stack.WindowManagement.WinApi {
    using System;
    using System.Linq;

    public enum HResult: uint {
        RPC_E_CANTCALLOUT_ININPUTSYNCCALL = 0x8001010D,
        TYPE_E_ELEMENTNOTFOUND = 0x8002802B,
        ERROR_INVALID_CURSOR_HANDLE = 0x8007057A,
    }

    public static class HResultExtensions {
        public static bool EqualsCode(this HResult result, int code) => result == (HResult)code;

        public static bool Match(this Exception ex, params HResult[] hResult) {
            return hResult.Select(x => (uint)x).Any(x => ((uint)ex.HResult) == x);
        }
    }
}