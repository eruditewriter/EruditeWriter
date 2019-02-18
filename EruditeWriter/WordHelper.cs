using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Word = Microsoft.Office.Interop.Word;

namespace EruditeWriter
{
    public class MSWordApp : HwndHost
    {
        internal const int
            WS_BORDER = 0x00800000,
            WS_CAPTION = 0x00C00000,
            WS_CHILD = 0x40000000,
            WS_CLIPSIBLINGS = 0x04000000,
            WS_CLIPCHILDREN = 0x02000000,
            WS_MAXIMIZEBOX = 0x00010000,
            WS_MINIMIZEBOX = 0x00020000,
            WS_VISIBLE = 0x10000000,
            GWL_STYLE = -16,
            SWP_FRAMECHANGED = 0x0020,
            SWP_NO_MOVE = 0x0002,

            LbsNotify = 0x00000001,
            HostId = 0x00000002,
            ListboxId = 0x00000001,
            WsVscroll = 0x00200000;

        private EWApp thisApp = (EWApp)Application.Current;

        public Word.Application wordApp = null;
        public Word.Document wordDoc = null;
        public Window winMask = null;
        public WindowInteropHelper winMaskWIH = null;
        public int wordDocNumber = 0;

        /*public WordApp(double height, double width)
        {
             _hostHeight = height;
             _hostWidth = width;
        }*/
        public MSWordApp()
        {

        }

        public IntPtr HwndWordApp { get; private set; }

        protected override HandleRef BuildWindowCore(HandleRef hwndParent)
        {
            //get the coordinates of our host element on the main screen
            Point pos = thisApp.mainWin.WordHostElement.PointToScreen(new Point(0, 0));
            //Point posmw = System.Windows.PresentationSource.FromVisual(thisApp.mainWin).CompositionTarget.TransformFromDevice.Transform(pos);
            //Point posmw = thisApp.mainWin.PointToScreen(new Point(0, 0));
            //var absolutePos = new Point(pos.X - posmw.X, pos.Y - posmw.Y);
            //get dpi scale factor
            double factor = System.Windows.PresentationSource.FromVisual(thisApp.mainWin).CompositionTarget.TransformToDevice.M11;
            //hieght and width need to be adjusted based on the dip as actual height and width are based on absolute pixel
            var h = thisApp.mainWin.WordHostElement.ActualHeight * factor;
            var w = thisApp.mainWin.WordHostElement.ActualWidth * factor;
            //set Word to be visible when opening the documents
            wordApp.Visible = true;
            //open the first document for the monograph
            wordDoc = wordApp.Documents.Open(thisApp.dirLocation + "\\Draft\\1.docx");
            //loop through the open word documents to get the window and then handle for our document
            for (int i = 1; i <= wordApp.Windows.Count; i++)
            { //we have open documents
                if (wordApp.Windows[i].Document.FullName == thisApp.dirLocation + "\\Draft\\1.docx")
                {
                    HwndWordApp = (IntPtr)wordApp.Windows[i].Hwnd;
                    wordDocNumber = i;
                    wordApp.Windows[i].WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                    break;
                }
            }
            //with Word handle we change style and set parent so we can properly host inside our app
            SetWindowLongPtr(HwndWordApp, GWL_STYLE, (IntPtr)(WS_BORDER | WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN));
            SetParent(HwndWordApp, hwndParent.Handle);
            //SetWindowPos(HwndWordApp, IntPtr.Zero, 0, 0, (int)w, (int)h, SWP_FRAMECHANGED | SWP_NO_MOVE);
            //create the mask to block use of sys_menu, File tab, and close button
            winMask = new WindowMask();
            winMaskWIH = new WindowInteropHelper(winMask);
            winMask.Owner = thisApp.mainWin;
            winMask.Show();
            SetWindowLongPtr(winMaskWIH.Handle, GWL_STYLE, (IntPtr)(WS_BORDER | WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN));
            SetWindowPos(winMaskWIH.Handle, IntPtr.Zero, (int)pos.X + 1, (int)pos.Y + 2, (int)w, (int)h, 0);

            this.InvalidateVisual();

            return new HandleRef(this, HwndWordApp);
        }

        protected override IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            handled = false;
            return IntPtr.Zero;
        }

        protected override void DestroyWindowCore(HandleRef hwnd)
        {
            DestroyWindow(hwnd.Handle);
        }

        //define external win32 imports
        static IntPtr GetWindowLongPtr(IntPtr hWnd, Int32 nIndex)
        {//GetWindowLongPtr on 32 bit calls GetWindowLong so setup for 32bit and 64bit calls
            if (IntPtr.Size == 4)
            {
                return GetWindowLongPtr32(hWnd, nIndex);
            }
            else
            {
                return GetWindowLongPtr64(hWnd, nIndex);
            }
        }

        static IntPtr SetWindowLongPtr(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong)
        {//SetWindowLongPtr on 32 bit calls GetWindowLong so setup for 32bit and 64bit calls
            if (IntPtr.Size == 4)
            {
                return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
            }
            else
            {
                return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            }
        }

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLong")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return", Justification = "This declaration is not used on 64-bit Windows.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2", Justification = "This declaration is not used on 64-bit Windows.")]
        private static extern IntPtr GetWindowLongPtr32(IntPtr hWnd, Int32 nIndex);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLongPtr")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist", Justification = "Entry point does exist on 64-bit Windows.")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, Int32 nIndex);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLong")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return", Justification = "This declaration is not used on 64-bit Windows.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2", Justification = "This declaration is not used on 64-bit Windows.")]
        private static extern IntPtr SetWindowLongPtr32(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLongPtr")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist", Justification = "Entry point does exist on 64-bit Windows.")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "DestroyWindow", CharSet = CharSet.Unicode)]
        internal static extern bool DestroyWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern int SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "MoveWindow")]
        static extern bool MoveWindow(
                IntPtr hWnd,
                int X,
                int Y,
                int nWidth,
                int nHeight,
                bool bRepaint
            );

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        static extern bool SetWindowPos(
               IntPtr hWnd,               // handle to window
               IntPtr hWndInsertAfter,    // placement-order handle
               int X,                  // horizontal position
               int Y,                  // vertical position
               int cx,                 // width
               int cy,                 // height
               uint uFlags             // window-positioning options
          );

    } //end WordApp
} //end namespace

    /* public class WordApp
     {
          [DllImport("user32.dll")]
          static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

          [DllImport("user32.dll")]
          static extern int SetParent(int hWndChild, int hWndNewParent);

          [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
          static extern bool SetWindowPos(
               int hWnd,               // handle to window
               int hWndInsertAfter,    // placement-order handle
               int X,                  // horizontal position
               int Y,                  // vertical position
               int cx,                 // width
               int cy,                 // height
               uint uFlags             // window-positioning options
          );

          [DllImport("user32.dll", EntryPoint = "MoveWindow")]
          static extern bool MoveWindow(
               int hWnd,
               int X,
               int Y,
               int nWidth,
               int nHeight,
               bool bRepaint
          );

          public struct RECT
          {
               public long left, top, right, bottom;
          }

          [DllImport("user32.dll", EntryPoint = "GetClientRect")]
          static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

          static IntPtr GetWindowLongPtr(IntPtr hWnd, Int32 nIndex)
          {//GetWindowLongPtr on 32 bit calls GetWindowLong so setup for 32bit and 64bit calls
               if (IntPtr.Size == 4)
               {
                    return GetWindowLongPtr32(hWnd, nIndex);
               }
               else
               {
                    return GetWindowLongPtr64(hWnd, nIndex);
               }
          }

          static IntPtr SetWindowLongPtr(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong)
          {//SetWindowLongPtr on 32 bit calls GetWindowLong so setup for 32bit and 64bit calls
               if (IntPtr.Size == 4)
               {
                    return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
               }
               else
               {
                    return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
               }
          }

          [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLong")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return", Justification = "This declaration is not used on 64-bit Windows.")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2", Justification = "This declaration is not used on 64-bit Windows.")]
          private static extern IntPtr GetWindowLongPtr32(IntPtr hWnd, Int32 nIndex);

          [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLongPtr")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist", Justification = "Entry point does exist on 64-bit Windows.")]
          private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, Int32 nIndex);

          [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLong")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return", Justification = "This declaration is not used on 64-bit Windows.")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2", Justification = "This declaration is not used on 64-bit Windows.")]
          private static extern IntPtr SetWindowLongPtr32(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong);

          [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLongPtr")]
          [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist", Justification = "Entry point does exist on 64-bit Windows.")]
          private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, Int32 nIndex, IntPtr dwNewLong);

          const int SWP_DRAWFRAME = 0x20;
          const int SWP_NOMOVE = 0x2;
          const int SWP_NOSIZE = 0x1;
          const int SWP_NOZORDER = 0x4;
          const int SWP_SHOWWINDOW = 0x40;
          const long WS_BORDER = 0x00800000L;
          const long WS_CHILD = 0x40000000L;
          const long WS_VISIBLE = 0x10000000L;
          const long WS_MAXIMIZE = 0x01000000L;

          public Word.Application wordApp = null;
          public Word.Document wordDoc = null;

          public int wordhWnd = 0;
          public IntPtr hWnd, coverhWnd;

          public Window coverwin = null;

          public int StartWord()
          {
               App thisApp = (App)Application.Current;
          Window window = Window.GetWindow(thisApp.mainWin);
               WindowInteropHelper wih = new WindowInteropHelper(window);
               hWnd = wih.Handle;
               wordApp = new Word.Application();
               Word.Application wordAppOrig = null;
               try
               {
                    wordAppOrig = (Word.Application)Marshal.GetActiveObject("Word.Application");
                    for (int i = 1; i <= wordAppOrig.Windows.Count; i++)
                    { //we have open documents
                         if (wordAppOrig.Windows[i].Document.FullName == thisApp.dirLocation + "\\Draft\\1.docx")
                         {
                              wordApp = wordAppOrig;
                              wordDoc = wordApp.Windows[i].Document;
                              wordhWnd = wordApp.Windows[i].Hwnd;
                              wordAppOrig = null;
                              break;
                         }
                    }
               }
               catch
               {
                    //no documents open
               }
               finally
               {
                    if (wordAppOrig != null)
                    {
                         wordAppOrig = null;
                    }
               }
               if (wordDoc == null)
               {
                    wordDoc = wordApp.Documents.Open(thisApp.dirLocation + "\\Draft\\1.docx");
                    for (int i = 1; i <= wordApp.Windows.Count; i++)
                    { //we have open documents
                         if (wordApp.Windows[i].Document.FullName == thisApp.dirLocation + "\\Draft\\1.docx")
                         {
                              wordhWnd = wordApp.Windows[i].Hwnd;
                              break;
                         }
                    }
               }
               var h = thisApp.mainWin.layer0.ActualHeight;
               var w = thisApp.mainWin.layer0.ActualWidth;
               Point pos = thisApp.mainWin.layer0.PointToScreen(new Point(0, 0));
               Point posmw = thisApp.mainWin.PointToScreen(new Point(0, 0));
               var absolutePos = new Point(pos.X - posmw.X, pos.Y - posmw.Y);
               //get dpi scale factor
               double factor = System.Windows.PresentationSource.FromVisual(thisApp.mainWin).CompositionTarget.TransformToDevice.M11;
               h = h * factor;
               w = w * factor;
               var x = absolutePos.X;
               var y = absolutePos.Y;
               wordApp.Visible = true;
               //wordDoc.Activate();
               IntPtr style = GetWindowLongPtr((IntPtr)wordhWnd, -16);
               SetWindowLongPtr((IntPtr)wordhWnd, -16, (IntPtr)(WS_BORDER | WS_CHILD | WS_VISIBLE));
               SetParent(wordhWnd, wih.Handle.ToInt32());
               MoveWindow(wordhWnd, (int)x, (int)y, (int)w, (int)h, true);
               //SetWindowPos(wordhWnd, wih.Handle.ToInt32(), (int)x, (int)y, (int)w, (int)h, 0);
               coverwin = new Cover();
               //HwndSource hwnd = (HwndSource)HwndSource.FromVisual(coverwin);
               coverwin.Show();
               WindowInteropHelper wih2 = new WindowInteropHelper(coverwin);
               coverhWnd = wih2.Handle;
               SetParent(coverhWnd.ToInt32(), wih.Handle.ToInt32());
               MoveWindow(coverhWnd.ToInt32(), (int)x, (int)y, 75, 64, true);

               return 0;
          }

          public int StartMSWord()
          {
               App thisApp = (App)Application.Current;
               WordHost wordWin = new WordHost(thisApp.mainWin.ControlHostElement.ActualHeight, thisApp.mainWin.ControlHostElement.ActualWidth);
               thisApp.mainWin.ControlHostElement.Child = wordWin;
               return 0;
          }
     }*/