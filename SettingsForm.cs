using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using App.Utils;

public class SettingsForm : Form
{

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);
    const int SW_RESTORE = 9;

    private readonly TextBox titleBox;
    private readonly TextBox classBox;
    private readonly Button windowCheckButton;
    public string WindowTitle => titleBox.Text;
    public string WindowClass => classBox.Text;


    public SettingsForm()
    {
        Text = "設定";
        Width = 480;
        Height = 520;
        StartPosition = FormStartPosition.CenterParent;
        Load += SettingsForm_Load;

        Label label = new()
        {
            Text = "貼付対象設定:",
            Left = 20,
            Top = 20,
            Width = 100
        };

        Label lblTitle = new() { Text = "タイトル:", Left = 30, Top = 50, Width = 50 };

        titleBox = new TextBox()
        {
            Left = 80,
            Top = 50,
            Width = 200,
            Height = 20,
            Multiline = true,
        };

        Label lblClass = new() { Text = "クラス:", Left = 30, Top = 80, Width = 50 };

        classBox = new TextBox()
        {
            Left = 80,
            Top = 80,
            Width = 200,
            Height = 20,
            Multiline = true,
        };

        windowCheckButton = new Button()
        {
            Text = "チェック",
            Left = 200,
            Top = 110,
            Width = 90,
            Height = 20,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        windowCheckButton.FlatAppearance.BorderColor = Color.Black;
        windowCheckButton.FlatAppearance.BorderSize = 1;
        windowCheckButton.Click += (s, e) => CheckWindow();


        Button saveButton = new()
        {
            Text = "保存",
            Left = 220,
            Top = 200,
            Width = 90,
            Height = 20,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        saveButton.FlatAppearance.BorderColor = Color.Black;
        saveButton.FlatAppearance.BorderSize = 1;

        saveButton.Click += (sender, e) =>
        {
            var settings = AppSettings.Get();
            settings.WindowTitle = titleBox.Text;
            settings.WindowClass = classBox.Text;
            AppSettings.Save();
            DialogResult = DialogResult.OK;
        };

        Controls.Add(label);
        Controls.Add(saveButton);

        Controls.Add(lblTitle);
        Controls.Add(titleBox);
        Controls.Add(lblClass);
        Controls.Add(classBox);
        Controls.Add(windowCheckButton);
    }
    private void CheckWindow()
    {
        IntPtr hWnd = WindowHelpers.FindWindowByClassAndTitle(classBox.Text, titleBox.Text);
        if (hWnd == IntPtr.Zero)
        {
            MessageBox.Show("貼付対象ウィンドウが見つかりません。まずウィンドウを開くまたはタイトルやクラス名を直してください。", "ウィンドウが見つかりません",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 貼付対象ウィンドウを復元してフォーカスする
        if (IsIconic(hWnd))
        {
            ShowWindow(hWnd, SW_RESTORE);
        }
        SetForegroundWindow(hWnd);
    }

    private void SettingsForm_Load(object sender, EventArgs e)
    {
        var settings = AppSettings.Get();
        titleBox.Text = settings.WindowTitle;
        classBox.Text = settings.WindowClass;
    }
}
