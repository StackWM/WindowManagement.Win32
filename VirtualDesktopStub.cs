﻿namespace LostTech.Stack.WindowManagement {
    using System;
    using System.Runtime.CompilerServices;
    using WindowsDesktop;

    class VirtualDesktopStub {
        static readonly bool virtualDesktopLoaded;

        static VirtualDesktopStub() {
            string currentAssemblyDirectory;
            try {
                currentAssemblyDirectory = System.Reflection.Assembly.GetExecutingAssembly().Location;
                currentAssemblyDirectory = System.IO.Path.GetDirectoryName(currentAssemblyDirectory);
            } catch {
                return;
            }

            string virtualDesktopPath = System.IO.Path.Combine(currentAssemblyDirectory, "VirtualDesktop.dll");
            try {
                System.Reflection.Assembly.LoadFile(virtualDesktopPath);
            } catch (System.IO.FileNotFoundException) { } catch (System.IO.DirectoryNotFoundException) { }

            virtualDesktopLoaded = true;
        }
        public static bool HasMinimalSupport => virtualDesktopLoaded && GetHasMinimalSupport();
        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool GetHasMinimalSupport() => VirtualDesktop.HasMinimalSupport;

        public static bool IsSupported => virtualDesktopLoaded && GetIsSupported();
        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool GetIsSupported() => VirtualDesktop.IsSupported;

        [MethodImpl(MethodImplOptions.NoInlining)]
        public static Guid? IdFromHwnd(IntPtr handle) => VirtualDesktop.IdFromHwnd(handle);

        [MethodImpl(MethodImplOptions.NoInlining)]
        public static bool IsCurrentVirtualDesktop(IntPtr handle) => VirtualDesktopHelper.IsCurrentVirtualDesktop(handle);

        [Obsolete]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static bool IsPinnedWindow(IntPtr handle) => VirtualDesktop.IsPinnedWindow(handle);
    }
}
