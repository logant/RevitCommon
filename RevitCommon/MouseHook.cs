//using System;
//using System.Diagnostics;
//using System.Runtime.InteropServices;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace RevitCommon
//{
//    /// <summary>
//    /// Class for intercepting low level Windows mouse hooks
//    /// </summary>
//    class MouseHook
//    {
//        /// <summary>
//        /// Internal callback processing function
//        /// </summary>
//        /// <param name="nCode"></param>
//        /// <param name="wParam"></param>
//        /// <param name="lParam"></param>
//        /// <returns></returns>
//        private delegate IntPtr MouseHookHandler(int nCode, IntPtr wParam, IntPtr lParam);
//        private MouseHookHandler hookHandler;

//        /// <summary>
//        /// Function to be called when defined event occurs
//        /// </summary>
//        /// <param name="mouseStruct"></param>
//        public delegate void MouseHookCallback(MSLLHOOKSTRUCT mouseStruct);

//        #region Events
//        public event MouseHookCallback MouseActivate;
//        public event MouseHookCallback NCHitTest;
//        public event MouseHookCallback NCMouseMove;
//        public event MouseHookCallback NCLeftButtonDown;
//        public event MouseHookCallback NCLeftButtonUp;
//        public event MouseHookCallback NCLeftButtonDbClick;
//        public event MouseHookCallback NCRightButtonDown;
//        public event MouseHookCallback NCRightButtonUp;
//        public event MouseHookCallback NCRightButtonDbClick;
//        public event MouseHookCallback NCMiddleButtonDown;
//        public event MouseHookCallback NCMiddleButtonUp;
//        public event MouseHookCallback NCMiddleButtonDbClick;
//        public event MouseHookCallback NCXButtonDown;
//        public event MouseHookCallback NCXButtonUp;
//        public event MouseHookCallback NCXButtonDbClick;
//        public event MouseHookCallback MouseMove;
//        public event MouseHookCallback LeftButtonDown;
//        public event MouseHookCallback LeftButtonUp;
//        public event MouseHookCallback LeftButtonDbClick;
//        public event MouseHookCallback RightButtonDown;
//        public event MouseHookCallback RightButtonUp;
//        public event MouseHookCallback RightButtonDbClick;
//        public event MouseHookCallback MiddleButtonDown;
//        public event MouseHookCallback MiddleButtonUp;
//        public event MouseHookCallback MiddleButtonDbClick;
//        public event MouseHookCallback MouseWheel;
//        public event MouseHookCallback XButtonDown;
//        public event MouseHookCallback XButtonUp;
//        public event MouseHookCallback XButtonDbClick;
//        public event MouseHookCallback CaptureChanged;
//        public event MouseHookCallback NCMouseHover;
//        public event MouseHookCallback MouseHover;
//        public event MouseHookCallback NCMouseLeave;
//        #endregion

//        private IntPtr hookID = IntPtr.Zero;

//        public void Install()
//        {
//            hookHander = HookFunc;
//            hookID = SetHook(hookHandler);
//        }

//        public void Uninstall()
//        {
//            if (hookID == IntPtr.Zero)
//                return;
//            UnhookWindowsHookEx(hookID);
//            hookID = IntPtr.Zero;
//        }

//        /// <summary>
//        /// Destructor. Unhook current hook
//        /// </summary>
//        ~MouseHook()
//        {
//            Uninstall();
//        }

//        /// <summary>
//        /// Sets hook and assigns its ID for tracking
//        /// </summary>
//        /// <param name="proc">Internal callback function</param>
//        /// <returns> Hook ID </returns>
//        private IntPtr SetHook(MouseHookHandler proc)
//        {
//            using (ProcessModule module = Process.GetCurrentProcess().MainModule)
//                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(module.ModuleName), 0);
//        }

//        /// <summary>
//        /// Callback function
//        /// </summary>
//        /// <param name="nCode"></param>
//        /// <param name="wParam"></param>
//        /// <param name="lParam"></param>
//        /// <returns></returns>
//        private IntPtr HookFunc(int nCode, IntPtr wParam, IntPtr lParam)
//        {
//            // parse system messages
//            if (nCode >= 0)
//            {
//                if(MouseMessages.WM_MOUSEACTIVATE == (MouseMessages)wParam)
//                    if (MouseActivate != null)
//                        MouseActivate((MSLLHOOKSTRUCT) Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT)));
//            }

//            return CallNextHookEx(hookID, nCode, wParam, lParam);
//        }


//        [StructLayout(LayoutKind.Sequential)]
//        public struct POINT
//        {
//            public int x;
//            public int y;
//        }

//        [StructLayout(LayoutKind.Sequential)]
//        public struct MSLLHOOKSTRUCT
//        {
//            public POINT pt;
//            public uint mouseData;
//            public uint flags;
//            public uint time;
//            public IntPtr dwExtraInfo;
//        }

//        private enum MouseMessages
//        {
//            WM_MOUSEACTIVATE = 0x0021,      // Sent when the cursor is in an inactive window and a mouse button is pressed.
//            WM_NCHITTEST = 0x0084,          // Sent to a window in order to determine what part of the window corresponds to a particular screen coordinate
//            WM_NCMOUSEMOVE = 0x00A0,        // Posted when cursor is moved within a nonclient area of a window
//            WM_NCLBUTTONDOWN = 0x00A1,      // Posted when left mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCLBUTTONUP = 0x00A2,        // Posted when left mouse button is released while cursor is in a nonclient area of a window
//            WM_NCLBUTTONDBCLK = 0x00A3,     // Posted when double-clicking the left mouse button while the cursor is in a nonclient area of a window
//            WM_NCRBUTTONDOWN = 0x00A4,      // Posted when right mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCRBUTTONUP = 0x00A5,        // Posted when right mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCRBUTTONDBCLK = 0x00A6,     // Posted when double-clicking the right mouse button while the cursor is in a nonclient area of a window
//            WM_NCMBUTTONDOWN = 0x00A7,      // Posted when middle mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCMBUTTONUP = 0x00A8,        // Posted when middle mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCMBUTTONDBCLK = 0x00A9,     // Posted when double-clicking the middle mouse button while the cursor is in a nonclient area of a window
//            WM_NCXBUTTONDOWN = 0x00AB,      // Posted when the first or second X mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCXBUTTONUP = 0x00AC,        // Posted when the first or second X  mouse button is pressed while cursor is in a nonclient area of a window
//            WM_NCXBUTTONDBCLK = 0x00AD,     // Posted when double-clicking the first or second X mouse button while the cursor is in a nonclient area of a window
//            WM_MOUSEMOVE = 0x0200,          // Posted when the cursor moves
//            WM_LBUTTONDOWN = 0x0201,        // Posted when left mouse button is pressed while cursor is in the client area of a window
//            WM_LBUTTONUP = 0x0202,          // Posted when releasing left mouse button while cursor is in the client area of a window
//            WM_LBUTTONDBCLK = 0x0203,       // Posted when double-clicking left mouse button while cursor is in the client area of a window
//            WM_RBUTTONDOWN = 0x0204,        // Posted when right mouse button is pressed while cursor is in the client area of a window
//            WM_RBUTTONUP = 0x0205,          // Posted when releasing right mouse button while cursor is in the client area of a window
//            WM_RBUTTONDBCLK = 0x0206,       // Posted when double-clicking the right mouse button while cursor is in the client area of a window
//            WM_MBUTTONDOWN = 0x0207,        // Posted when right mouse button is pressed while cursor is in the client area of a window
//            WM_MBUTTONUP = 0x0208,          // Posted when releasing right mouse button while cursor is in the client area of a window
//            WM_MBUTTONDBCLK = 0x0209,       // Posted when double-clicking the middle mouse button while cursor is in the client area of a window
//            WM_MOUSEWHEEL = 0x020A,         // Posted when the mouse wheel is rotated
//            WM_XBUTTONDOWN = 0x020B,        // Posted when the first or second X button is pressed while in client area of a window.
//            WM_XBUTTONUP = 0x020C,          // Posted when the user releases the first or second X button
//            WM_XBUTTONDBCLK = 0x020D,       // Posted when the user double-clicks the first or second X button 
//            WM_CAPTURECHANGED = 0x0215,     // Message when losing the mouse capture
//            WM_NCMOUSEHOVER = 0x02A0,       // Posted when a cursor hovers over the nonclient area of a window for a predefined period of time.
//            WM_MOUSEHOVER = 0x02A1,         // Posted when a cursor hovers over the client area of a window for a predefined period of time.
//            WM_NCMOUSELEAVE = 0x02A2        // Posted when the cursor leaves the nonclient area of a window  in a prior call. 
//        }

//        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//        private static extern IntPtr SetWindowsHookEx(int idHook, MouseHookHandler lpfn, IntPtr hMod, uint dwThreadId);

//        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//        [return: MarshalAs(UnmanagedType.Bool)]
//        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

//        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

//        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//        private static extern IntPtr GetModuleHandle(string lpModuleName);
//    }
//}
