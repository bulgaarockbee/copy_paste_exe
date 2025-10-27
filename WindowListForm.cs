using System;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class WindowListForm : Form
{
    public string SelectedTitle { get; private set; } = "";
    public string SelectedClass { get; private set; } = "";

    private readonly ListBox listBox;
    private readonly Button cancelButton;
    private readonly Button okButton;
    private List<WindowInfo> windows;

    public WindowListForm()
    {
        Text = "ウィンドウ選択";
        Width = 600;
        Height = 400;

        listBox = new ListBox { Left = 10, Top = 10, Width = 560, Height = 300 };
        cancelButton = new Button { Left = 10, Top = 320, Width = 100, Text = "キャンセル" };
        okButton = new Button { Left = 120, Top = 320, Width = 100, Text = "選択" };
        okButton.Click += cancelButton_Click;
        okButton.Click += OkButton_Click;

        Controls.Add(listBox);
        Controls.Add(cancelButton);
        Controls.Add(okButton);

        Load += WindowListForm_Load;
    }

    private void WindowListForm_Load(object sender, EventArgs e)
    {
        windows = [];

        EnumWindows((hWnd, lParam) =>
        {
            if (IsWindowVisible(hWnd))
            {
                StringBuilder title = new StringBuilder(256);
                GetWindowText(hWnd, title, title.Capacity);

                StringBuilder className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);

                if (!string.IsNullOrWhiteSpace(title.ToString()))
                {
                    var win = new WindowInfo
                    {
                        Handle = hWnd,
                        Title = title.ToString(),
                        ClassName = className.ToString()
                    };
                    windows.Add(win);
                    listBox.Items.Add($"{win.Title} ({win.ClassName})");
                }
            }
            return true;
        }, IntPtr.Zero);
    }

    private void cancelButton_Click(object sender, EventArgs e)
    {
        Close();
    }

    private void OkButton_Click(object sender, EventArgs e)
    {
        if (listBox.SelectedIndex >= 0)
        {
            var win = windows[listBox.SelectedIndex];
            SelectedTitle = win.Title;
            SelectedClass = win.ClassName;
            DialogResult = DialogResult.OK;
            Close();
        }
        else
        {
            MessageBox.Show("リストからウィンドウを選択してください");
        }
    }

    private class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; } = "";
        public string ClassName { get; set; } = "";
    }

    // P/Invoke
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}
