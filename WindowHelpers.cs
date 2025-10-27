using System;
using System.Runtime.InteropServices;

namespace App.Utils
{
    public static class WindowHelpers
    {

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public static IntPtr FindWindowByClassAndTitle(string classNameFilter, string titleFilter)
        {
            IntPtr foundWindow = IntPtr.Zero;

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                var className = new System.Text.StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);
                string cls = className.ToString();

                var windowText = new System.Text.StringBuilder(256);
                GetWindowText(hWnd, windowText, windowText.Capacity);
                string title = windowText.ToString();

                // クラスとタイトルの両方に一致（大文字と小文字を区別しない、部分一致）
                if (!string.IsNullOrEmpty(cls) && !string.IsNullOrEmpty(title) &&
                    cls.IndexOf(classNameFilter, StringComparison.OrdinalIgnoreCase) >= 0 &&
                    title.IndexOf(titleFilter, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundWindow = hWnd;
                    return false; // 列挙を停止する
                }

                return true; // 検索を続ける
            }, IntPtr.Zero);

            return foundWindow;
        }

    }
}
