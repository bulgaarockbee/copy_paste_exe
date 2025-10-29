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
    private readonly NumericUpDown pressCountBox;
    private readonly NumericUpDown delayBox;
    private readonly Button windowCheckButton;
    private readonly Button windowSelectButton;
    private RadioButton rbOption1;
    private RadioButton rbOption2;
    private RadioButton rbOption3;
    private RadioButton rbOption4;
    public string WindowTitle => titleBox.Text;
    public string WindowClass => classBox.Text;
    public string ButtonType => rbOption1.Checked ? rbOption1.Text :
                             rbOption2.Checked ? rbOption2.Text :
                             rbOption3.Checked ? rbOption3.Text :
                             rbOption4.Checked ? rbOption4.Text :
                             string.Empty;
    public string PressCount => pressCountBox.Text;
    public string Delay => delayBox.Text;


    public SettingsForm()
    {
        Text = "設定";
        Width = 480;
        Height = 400;
        StartPosition = FormStartPosition.CenterParent;
        Load += SettingsForm_Load;

        Label lblWindowSelectSetting = new()
        {
            Text = "貼付対象設定",
            Left = 20,
            Top = 20,
            Width = 100
        };

        Label lblTitle = new() { Text = "タイトル：", Left = 30, Top = 50, Width = 60 };

        titleBox = new TextBox()
        {
            Left = 90,
            Top = 50,
            Width = 350,
            Height = 20,
            Multiline = true,
        };

        Label lblClass = new() { Text = "クラス：", Left = 30, Top = 80, Width = 60 };

        classBox = new TextBox()
        {
            Left = 90,
            Top = 80,
            Width = 350,
            Height = 20,
            Multiline = true,
        };

        windowSelectButton = new Button()
        {
            Text = "選択",
            Left = 140,
            Top = 110,
            Width = 90,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        windowSelectButton.FlatAppearance.BorderColor = Color.Black;
        windowSelectButton.FlatAppearance.BorderSize = 1;
        windowSelectButton.Click += WindowSelectButton_Click;

        windowCheckButton = new Button()
        {
            Text = "チェック",
            Left = 240,
            Top = 110,
            Width = 90,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        windowCheckButton.FlatAppearance.BorderColor = Color.Black;
        windowCheckButton.FlatAppearance.BorderSize = 1;
        windowCheckButton.Click += (s, e) => CheckWindow();

        Label lblCursorMovementSetting = new()
        {
            Text = "項目遷移設定",
            Left = 20,
            Top = 160,
            Width = 100
        };

        Label lblButtonType = new() { Text = "種類：", Left = 30, Top = 190, Width = 50 };

        rbOption1 = new RadioButton() { Text = "Shift", Left = 80, Top = 190, Width = 50 };
        rbOption2 = new RadioButton() { Text = "Ctrl", Left = 130, Top = 190, Width = 50 };
        rbOption3 = new RadioButton() { Text = "Alt", Left = 180, Top = 190, Width = 50 };
        rbOption4 = new RadioButton() { Text = "Tab", Left = 230, Top = 190, Width = 50 };

        Label lblPressCount = new() { Text = "回数：", Left = 30, Top = 220, Width = 50 };

        pressCountBox = new NumericUpDown()
        {
            Left = 80,
            Top = 220,
            Width = 50,
            Height = 20,
            Maximum = 10
        };

        Label lblPressCountUnit = new() { Text = "回", Left = 130, Top = 220, Width = 20 };

        Label lblDelay = new() { Text = "待機：", Left = 30, Top = 250, Width = 50 };

        delayBox = new NumericUpDown()
        {
            Left = 80,
            Top = 250,
            Width = 50,
            Height = 20,
            Maximum = 10000
        };

        Label lblDelayUnit = new() { Text = "ms", Left = 130, Top = 250, Width = 20 };

        Button saveButton = new()
        {
            Text = "保存",
            Left = 190,
            Top = 300,
            Width = 90,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        saveButton.FlatAppearance.BorderColor = Color.Black;
        saveButton.FlatAppearance.BorderSize = 1;

        saveButton.Click += (sender, e) =>
        {
            var settings = AppSettings.Get();
            settings.CurrentWindowTitle = titleBox.Text;
            settings.CurrentWindowClass = classBox.Text;
            settings.CurrentPressCount = pressCountBox.Text;
            settings.CurrentDelay = delayBox.Text;
            settings.CurrentButtonType = rbOption1.Checked ? rbOption1.Text :
                             rbOption2.Checked ? rbOption2.Text :
                             rbOption3.Checked ? rbOption3.Text :
                             rbOption4.Checked ? rbOption4.Text :
                             string.Empty;

            AppSettings.Save();
            DialogResult = DialogResult.OK;
        };

        Controls.Add(lblWindowSelectSetting);
        Controls.Add(lblTitle);
        Controls.Add(titleBox);
        Controls.Add(lblClass);
        Controls.Add(classBox);
        Controls.Add(windowSelectButton);
        Controls.Add(windowCheckButton);
        Controls.Add(lblCursorMovementSetting);
        Controls.Add(lblButtonType);
        Controls.Add(rbOption1);
        Controls.Add(rbOption2);
        Controls.Add(rbOption3);
        Controls.Add(rbOption4);
        Controls.Add(lblPressCount);
        Controls.Add(pressCountBox);
        Controls.Add(lblPressCountUnit);
        Controls.Add(lblDelay);
        Controls.Add(delayBox);
        Controls.Add(lblDelayUnit);
        Controls.Add(saveButton);
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

    private void WindowSelectButton_Click(object sender, EventArgs e)
    {
        using var dialog = new WindowListForm();
        if (dialog.ShowDialog() == DialogResult.OK)
        {
            titleBox.Text = dialog.SelectedTitle;
            classBox.Text = dialog.SelectedClass;
        }
    }

    private void SettingsForm_Load(object sender, EventArgs e)
    {
        var settings = AppSettings.Get();
        titleBox.Text = settings.CurrentWindowTitle;
        classBox.Text = settings.CurrentWindowClass;
        pressCountBox.Text = string.IsNullOrEmpty(settings.CurrentPressCount)
            ? "1"
            : settings.CurrentPressCount;
        delayBox.Text = string.IsNullOrEmpty(settings.CurrentDelay)
            ? "200"
            : settings.CurrentDelay;
        switch (settings.CurrentButtonType)
        {
            case var s when s == rbOption1.Text:
                rbOption1.Checked = true;
                break;
            case var s when s == rbOption2.Text:
                rbOption2.Checked = true;
                break;
            case var s when s == rbOption3.Text:
                rbOption3.Checked = true;
                break;
            case var s when s == rbOption4.Text:
                rbOption4.Checked = true;
                break;
            default:
                // タブをチェック
                rbOption4.Checked = true;
                break;
        }

    }
}
