#nullable enable

namespace LostTech.Stack.WindowManagement {
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Threading.Tasks;
    using JetBrains.Annotations;
    using LostTech.Stack.WindowManagement.WinApi;
    using PInvoke;
    using WindowsDesktop.Interop;
    using static PInvoke.User32;
    using Win32Exception = System.ComponentModel.Win32Exception;
    using Rect = System.Drawing.RectangleF;

    [DebuggerDisplay("{" + nameof(Title) + "}")]
    public sealed class Win32Window: IAppWindow, IEquatable<Win32Window> {
        readonly Lazy<bool> excludeFromMargin;
        public IntPtr Handle { get; }
        public bool SuppressSystemMargin { get; set; }

        public Win32Window(IntPtr handle, bool suppressSystemMargin) {
            this.Handle = handle;
            this.SuppressSystemMargin = suppressSystemMargin;
            this.excludeFromMargin = new Lazy<bool>(this.GetExcludeFromMargin);
        }

        const bool RepaintOnMove = true;
        public Task Move(Rect targetBounds) => Task.Run(async () => {
            var windowPlacement = WINDOWPLACEMENT.Create();
            if (GetWindowPlacement(this.Handle, ref windowPlacement) &&
                windowPlacement.showCmd.HasFlag(WindowShowStyle.SW_MAXIMIZE)) {
                ShowWindow(this.Handle, WindowShowStyle.SW_RESTORE);
            }

            if (this.SuppressSystemMargin && !this.excludeFromMargin.Value) {
                RECT systemMargin = GetSystemMargin(this.Handle);
#if DEBUG
                Debug.WriteLine($"{this.Title} compensating system margin {systemMargin.left},{systemMargin.top},{systemMargin.right},{systemMargin.bottom}");
#endif
                targetBounds.X -= systemMargin.left;
                targetBounds.Y -= systemMargin.top;
                targetBounds.Width = Math.Max(0, targetBounds.Width + systemMargin.left + systemMargin.right);
                targetBounds.Height = Math.Max(0, targetBounds.Height + systemMargin.top + systemMargin.bottom);
            }

            if (!MoveWindow(this.Handle, (int)targetBounds.Left, (int)targetBounds.Top,
                (int)targetBounds.Width, (int)targetBounds.Height, bRepaint: RepaintOnMove)) {
                var exception = this.GetLastError();
                if (exception is Win32Exception win32Exception
                    && win32Exception.NativeErrorCode == (int)WinApiErrorCode.ERROR_ACCESS_DENIED)
                    throw new UnauthorizedAccessException("Not enough privileges to move window", inner: exception);
                else
                    throw exception;
            } else {
                // TODO: option to not activate on move
                await Task.Yield();
#if DEBUG
                Debug.WriteLine($"{this.Title} final rect: {targetBounds}");
#endif
                MoveWindow(this.Handle, (int)targetBounds.Left, (int)targetBounds.Top, (int)targetBounds.Width,
                    (int)targetBounds.Height, bRepaint: RepaintOnMove);
            }
        });

        public Task<bool?> Close() => Task.Run(() => {
            IntPtr result = SendMessage(this.Handle, WindowMessage.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return (bool?)null;
        });

        public bool CanMove =>
            PostMessage(this.Handle, WindowMessage.WM_USER, IntPtr.Zero, IntPtr.Zero)
            || Marshal.GetLastWin32Error() != (int)WinApiErrorCode.ERROR_ACCESS_DENIED;

        /// <summary>
        /// Non-WPF coordinates
        /// </summary>
        public Rect Bounds {
            get {
                if (!Win32.GetWindowInfo(this.Handle, out var info))
                    throw this.GetLastError();

                var bounds = new Rect(info.rcWindow.left, info.rcWindow.top,
                    info.rcWindow.right - info.rcWindow.left,
                    info.rcWindow.bottom - info.rcWindow.top);

                if (this.SuppressSystemMargin && !this.excludeFromMargin.Value) {
                    RECT systemMargin = GetSystemMargin(this.Handle);
                    bounds.X += systemMargin.left;
                    bounds.Y += systemMargin.top;
                    bounds.Width = Math.Max(0, bounds.Width - (systemMargin.left + systemMargin.right));
                    bounds.Height = Math.Max(0, bounds.Height - (systemMargin.top + systemMargin.bottom));
                }

                return bounds;
            }
        }

        public Task<Rect> GetBounds() => Task.Run(() => this.Bounds);

        public Task<Rect> GetClientBounds() => Task.Run(() => {
            DwmGetWindowAttribute(this.Handle,
                DwmApi.DWMWINDOWATTRIBUTE.DWMWA_CAPTION_BUTTON_BOUNDS, out RECT buttonBounds,
                Marshal.SizeOf<RECT>());

            var bounds = this.Bounds;
            if (buttonBounds.right == 0)
                return bounds;

            float titlebarLeft = bounds.Width - buttonBounds.right;
            float titlebarRight = buttonBounds.bottom;
            bounds.X += titlebarLeft;
            bounds.Y += titlebarRight;
            bounds.Width -= titlebarLeft;
            bounds.Height -= titlebarRight;
            return bounds;
        });

        public Task<IntPtr> GetIcon(int dpi = 96) => Task.Run(async () => {
            IntPtr hIcon = SendMessage(this.Handle, WindowMessage.WM_GETICON, (IntPtr)2 /*small*/, (IntPtr)dpi);
            if (hIcon != IntPtr.Zero) return hIcon;

            hIcon = GetClassLong(this.Handle, ClassLong.GCLP_HICONSM);
            if (hIcon != IntPtr.Zero) return hIcon;

            GetWindowThreadProcessId(this.Handle, lpdwProcessId: out int processID);
            if (processID == 0)
                return hIcon;

            try {
                string? exePath = GetExecutablePathAboveVista(processID);
                if (exePath?.EndsWith("ApplicationFrameHost.exe", StringComparison.OrdinalIgnoreCase) == true) {
                    Win32Window? uwpWindow = null;
                    this.ForEachChild(child => {
                        int threadID = GetWindowThreadProcessId(child.Handle, out int childProcessID);
                        if (threadID == 0 || childProcessID == processID)
                            return true;

                        uwpWindow = child;
                        return false;
                    });

                    if (uwpWindow != null)
                        return await uwpWindow.GetIcon(dpi).ConfigureAwait(false);
                }
                using var icon = Icon.ExtractAssociatedIcon(exePath);
                return icon.Handle;
            } catch (ArgumentException) { } catch (InvalidOperationException) { } catch (Exception e) {
                Debug.WriteLine(e);
                return hIcon;
            }

            return hIcon;
        });

        string? lastTitle;
        public string? Title {
            get {
                try {
                    return this.lastTitle = GetWindowText(this.Handle);
                } catch (PInvoke.Win32Exception) {
                    return null;
                }
            }
        }

        public string? Class {
            get {
                try {
                    return GetClassName(this.Handle);
                } catch (PInvoke.Win32Exception) {
                    return null;
                }
            }
        }

        public bool IsMinimized => IsIconic(this.Handle);
        public bool IsVisible {
            get {
                if (!IsWindowVisible(this.Handle))
                    return false;
                if (!VirtualDesktopStub.HasMinimalSupport)
                    return true;
                if (this.IsPopup)
                    return this.IsOnCurrentDesktop;

                Guid? desktopId = null;

                var timer = Stopwatch.StartNew();
                Exception? e = null;
                while (timer.Elapsed < this.ShellUnresposivenessTimeout) {
                    try {
                        desktopId = VirtualDesktopStub.IdFromHwnd(this.Handle);
                        e = null;
                        break;
                    } catch (COMException ex) when (ex.Match(WindowsDesktop.Interop.HResult.RPC_E_CANTCALLOUT_ININPUTSYNCCALL)) {
                        e = ex;
                        var async = Task.Run(() => desktopId = VirtualDesktopStub.IdFromHwnd(this.Handle));
                        var waitFor = this.ShellUnresposivenessTimeout - timer.Elapsed;
                        if (waitFor.Ticks < 0)
                            break;
                        try {
                            if (async.Wait(waitFor)) {
                                e = null;
                                break;
                            }
                        } catch (AggregateException asyncEx) {
                            e = asyncEx;
                            break;
                        }
                    } catch (COMException ex) {
                        e = ex;
                    }
                }

                if (desktopId == null) {
                    if (e != null)
                        throw new ShellUnresponsiveException(e);

                    this.Closed?.Invoke(this, EventArgs.Empty);
                    throw this.WindowNotFoundException_();
                }
                return desktopId != null && desktopId != Guid.Empty;
            }
        }

        public bool? IsCloaked {
            get {
                var result = DwmGetWindowAttribute(this.Handle, out Cloaking value);
                if (!result.Succeeded) return null;
                return value != Cloaking.None;
            }
        }

        public bool IsValid => IsWindow(this.Handle);
        public bool IsOnCurrentDesktop {
            get {
                if (!VirtualDesktopStub.HasMinimalSupport)
                    return true;

                try {
                    return VirtualDesktopStub.IsCurrentVirtualDesktop(this.Handle);
                } catch (COMException e)
                    when (WinApi.HResult.TYPE_E_ELEMENTNOTFOUND.EqualsCode(e.HResult)) {
                    this.Closed?.Invoke(this, EventArgs.Empty);
                    throw this.WindowNotFoundException_(innerException: e);
                } catch (COMException e) {
                    e.ReportAsWarning();
                    return true;
                } catch (Win32Exception e) {
                    e.ReportAsWarning();
                    return true;
                } catch (ArgumentException e) {
                    e.ReportAsWarning();
                    return true;
                }
            }
        }

        public bool IsVisibleInAppSwitcher => Win32WindowFactory.DisplayInSwitchToList(this);

        public WindowStyles Styles => (WindowStyles)GetWindowLong(this.Handle, WindowLongIndexFlags.GWL_STYLE);
        public WindowStylesEx StylesEx => (WindowStylesEx)GetWindowLong(this.Handle, WindowLongIndexFlags.GWL_EXSTYLE);

        public bool IsResizable => this.Styles.HasFlag(WindowStyles.WS_SIZEFRAME);
        public bool IsPopup => this.Styles.HasFlag(WindowStyles.WS_POPUP);

        public Task<Exception?> Activate() {
            Exception? error = this.EnsureNotMinimized();
            return Task.FromResult(
                SetForegroundWindow(this.Handle) ? error : this.GetLastError());
        }

        public Task<Exception?> BringToFront() {
            Exception? issue = this.EnsureNotMinimized();
            if (!SetWindowPos(this.Handle, GetForegroundWindow(), 0, 0, 0, 0,
                              SetWindowPosFlags.SWP_NOMOVE | SetWindowPosFlags.SWP_NOACTIVATE |
                              SetWindowPosFlags.SWP_NOSIZE))
                issue = new Win32Exception();
            return Task.FromResult(issue);
        }

        public Task<Exception?> SendToBottom() {
            Exception? issue = this.EnsureNotMinimized();
            if (!SetWindowPos(this.Handle, HWND_BOTTOM, 0, 0, 0, 0,
                SetWindowPosFlags.SWP_NOMOVE | SetWindowPosFlags.SWP_NOACTIVATE |
                SetWindowPosFlags.SWP_NOSIZE))
                issue = new Win32Exception();
            return Task.FromResult(issue);
        }

        /// <summary>
        /// Unreliable, may fire multiple times
        /// </summary>
        public event EventHandler? Closed;

        [MustUseReturnValue]
        Exception? EnsureNotMinimized() {
            if (!IsIconic(this.Handle))
                return null;

            return ShowWindow(this.Handle, WindowShowStyle.SW_RESTORE) ? null : this.GetLastError();
        }

        public bool Equals(Win32Window? other) {
            if (other is null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return this.Handle.Equals(other.Handle);
        }

        public override bool Equals(object? obj) {
            if (obj is null)
                return false;
            if (ReferenceEquals(this, obj))
                return true;
            if (obj.GetType() != this.GetType())
                return false;
            return this.Equals((Win32Window)obj);
        }

        public Exception? ForEachChild(Func<Win32Window, bool> action) {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            var enumerator = new EnumChildProc((hwnd, _) => {
                var window = new Win32Window(hwnd, this.SuppressSystemMargin);
                return action(window);
            });
            bool done = EnumChildWindows(this.Handle, enumerator, IntPtr.Zero);
            GC.KeepAlive(enumerator);
            return done ? null : new Win32Exception();
        }

        bool GetExcludeFromMargin() {
            GetWindowThreadProcessId(this.Handle, lpdwProcessId: out int processID);
            if (processID == 0)
                return false;
            try {
                var process = Process.GetProcessById(processID);
                return process.ProcessName switch {
                    "explorer" or "Everything" => true,
                    _ => false,
                };
            } catch (ArgumentException) { } catch (InvalidOperationException) { } catch (Exception e) {
                Debug.WriteLine(e);
                return false;
            }
            return false;
        }

        static RECT GetSystemMargin(IntPtr handle) {
            PInvoke.HResult success = DwmGetWindowAttribute(handle, DwmApi.DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS,
                out RECT withMargin, Marshal.SizeOf<RECT>());
            if (!success.Succeeded) {
                Debug.WriteLine($"DwmGetWindowAttribute: {success.GetException()}");
                return new RECT();
            }

            if (!GetWindowRect(handle, out var noMargin)) {
                Debug.WriteLine($"GetWindowRect: {new Win32Exception()}");
                return new RECT();
            }

            return new RECT {
                left = withMargin.left - noMargin.left,
                top = withMargin.top - noMargin.top,
                right = noMargin.right - withMargin.right,
                bottom = noMargin.bottom - withMargin.bottom,
            };
        }

        public override int GetHashCode() => this.Handle.GetHashCode();

        [MustUseReturnValue]
        Exception GetLastError() {
            var exception = new Win32Exception();
            if (exception.NativeErrorCode == (int)WinApiErrorCode.ERROR_INVALID_WINDOW_HANDLE) {
                this.Closed?.Invoke(this, EventArgs.Empty);
                return new WindowNotFoundException(innerException: exception);
            } else
                return exception;
        }

        [DllImport("User32.dll", SetLastError = true)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.None)]
        static extern IntPtr GetClassLong(IntPtr hWnd, ClassLong parameter);
        [DllImport("User32.dll", SetLastError = true)]
        static extern bool EnumChildWindows(IntPtr parentHandle, EnumChildProc callback, IntPtr lParam);
        delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);
        static string? GetExecutablePathAboveVista(int ProcessId) {
            using (var hprocess = Kernel32.OpenProcess(0x1000, false, ProcessId)) {
                if (hprocess.DangerousGetHandle() != IntPtr.Zero) {
                    string imageName = Kernel32.QueryFullProcessImageName(hprocess);
                    if (!string.IsNullOrEmpty(imageName))
                        return imageName;
                }
            }
            return null;
        }

        [DllImport("Dwmapi.dll")]
        static extern PInvoke.HResult DwmGetWindowAttribute(IntPtr hwnd, DwmApi.DWMWINDOWATTRIBUTE attribute, out RECT value, int valueSize);
        [DllImport("Dwmapi.dll")]
        static extern PInvoke.HResult DwmGetWindowAttribute(IntPtr hwnd, DwmApi.DWMWINDOWATTRIBUTE attribute, out Cloaking value, int valueSize);

        static PInvoke.HResult DwmGetWindowAttribute(IntPtr hwnd, out Cloaking value)
            => DwmGetWindowAttribute(hwnd, DwmApi.DWMWINDOWATTRIBUTE.DWMWA_CLOAKED,
                                     out value,
                                     Marshal.SizeOf(typeof(Cloaking).GetEnumUnderlyingType()));

        [DllImport("User32.dll")]
        static extern bool IsIconic(IntPtr hwnd);

        // ReSharper disable InconsistentNaming
        static readonly IntPtr HWND_BOTTOM = new(1);
        static readonly IntPtr HWND_TOP = IntPtr.Zero;
        static readonly IntPtr HWND_NOTOPMOST = new(-2);
        // ReSharper restore InconsistentNaming

        public TimeSpan ShellUnresposivenessTimeout { get; set; } = TimeSpan.FromMilliseconds(300);

        enum ClassLong: int {
            GCLP_HICONSM = -34,
        }

        enum Cloaking: int {
            None = 0,
        }

        WindowNotFoundException WindowNotFoundException_(Exception? innerException = null)
            => new(innerException) {
                Data = {
                    { "LastTitle", this.lastTitle },
                    { "Handle", this.Handle },
                },
            };
    }
}
