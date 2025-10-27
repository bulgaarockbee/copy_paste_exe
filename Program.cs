using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using App.Utils;

class PasteToWindowForm : Form
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    const int SW_RESTORE = 9;

    private readonly TextBox subjectiveBox;
    private readonly TextBox objectiveBox;
    private readonly TextBox assessmentBox;
    private readonly TextBox planBox;
    private readonly TextBox voiceTextBox;
    private readonly Button pasteButton;
    private readonly Button insertVoiceText;
    private readonly Button aiSummarize;
    private string WindowTitle;
    private string WindowClass;
    private string ButtonType;
    private string PressCount;
    private string Delay;
    internal static readonly string[] items =
        [
            "SOAP(アシスト)"
        ];

    public PasteToWindowForm()
    {
        Text = "S.O.A.Pコピー";
        Width = 720;
        Height = 780;
        Load += PasteToWindowFormm_Load;

        Button settingsButton = new()
        {
            Text = "",
            Left = 650,
            Top = 10,
            Width = 30,
            Height = 30,
        };
        settingsButton.Image = Image.FromFile("settings.png");
        settingsButton.Image = new Bitmap(settingsButton.Image, new Size(30, 30));
        settingsButton.ImageAlign = ContentAlignment.MiddleCenter;

        insertVoiceText = new Button()
        {
            Text = "音声入力開始",
            Left = 300,
            Top = 10,
            Width = 100,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        insertVoiceText.FlatAppearance.BorderColor = Color.Black;
        insertVoiceText.FlatAppearance.BorderSize = 1;

        Label lblVoiceText = new() { Text = "音声書き起こし:", Left = 10, Top = 50, Width = 300 };
        voiceTextBox = new TextBox()
        {
            Left = 10,
            Top = 75,
            Width = 680,
            Height = 100,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        ComboBox colorDropdown = new()
        {
            Left = 200,
            Top = 185,
            Width = 150,
            DropDownStyle = ComboBoxStyle.DropDownList
        };

        colorDropdown.Items.AddRange(items);

        aiSummarize = new Button()
        {
            Text = "AI要約(デモ)",
            Left = 370,
            Top = 180,
            Width = 100,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        aiSummarize.FlatAppearance.BorderColor = Color.Black;
        aiSummarize.FlatAppearance.BorderSize = 1;

        Label lblS = new() { Text = "主観的情報（Subjective）:", Left = 10, Top = 210, Width = 300 };
        subjectiveBox = new TextBox()
        {
            Left = 10,
            Top = 235,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblO = new() { Text = "客観的情報（Objective）:", Left = 10, Top = 320, Width = 300 };
        objectiveBox = new TextBox()
        {
            Left = 10,
            Top = 345,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblA = new() { Text = "評価（Assessment）:", Left = 10, Top = 430, Width = 300 };
        assessmentBox = new TextBox()
        {
            Left = 10,
            Top = 455,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblP = new() { Text = "計画（Plan）:", Left = 10, Top = 540, Width = 300 };
        planBox = new TextBox()
        {
            Left = 10,
            Top = 565,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        pasteButton = new Button()
        {
            Text = "貼り付ける",
            Left = 300,
            Top = 660,
            Width = 100,
            Height = 30,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        pasteButton.FlatAppearance.BorderColor = Color.Black;
        pasteButton.FlatAppearance.BorderSize = 1;

        settingsButton.Click += (sender, e) =>
        {
            using SettingsForm settingsForm = new();
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                WindowTitle = settingsForm.WindowTitle;
                WindowClass = settingsForm.WindowClass;
                ButtonType = settingsForm.ButtonType;
                PressCount = settingsForm.PressCount;
                Delay = settingsForm.Delay;
            }
        };

        insertVoiceText.Click += (sender, e) =>
        {
            voiceTextBox.Text = "アムロジピン5mg朝食後を継続しメトホルミン250mg以上、朝夕食後に継続しロスバスタチン2.5mg1錠、夕食後、新規追加。血圧の変動はなく、血圧、糖尿病ともに特に問題ないとのことであり、起立時に軽度のめまいを自覚するとのこと。メトホルミンをまれに飲み忘れることがあるが内服継続中。新規追加薬に関して不安があり服薬指導希望。アムロジピンによる夜間性低血圧の可能性を考慮し、起立性低血圧に対する注意喚起と、服薬アドヒアランス向上のための一包化を提案。ロスバスタチン服用に伴う副作用（筋肉痛等）およびグレープフルーツジュース摂取によるCYP3A4阻害作用について説明し、回避を指導。今後も服薬アドヒアランスと副作用モニタリングのため定期的にフォロー予定。また、服薬管理方法について薬剤師と連携の上、継続的支援を実施する。コンプライアンス改善に向けて一包化を検討中。アムロジピン、メトホルミン、ロスバスタチン服用にて特段の副作用なし。";
        };

        aiSummarize.Click += (sender, e) =>
        {
            subjectiveBox.Text = "血圧低下、糖尿病悪化ともに特に問題なし起立性のめまいの訴え（軽度）メトホルミンをまれに飲み忘れることがある新規追加薬（ロスバスタチン）について不安の訴えあり";
            objectiveBox.Text = "アムロジピン 5mg 1錠 朝食後（継続）メトホルミン 250mg 朝夕食後（継続）ロスバスタチン 2.5mg 1錠 夕食後（新規追加）";
            assessmentBox.Text = "アムロジピンによる夜間性低血圧の可能性服薬コンプライアンス不良（メトホルミン飲み忘れ）ロスバスタチン新規開始時・副作用モニタリング必要服薬忘れ改善のための対策が必要";
            planBox.Text = "アムロジピンによるめまいについて説明、起立時注意を指導服薬コンプライアンス向上指導ロスバスタチン導入時（筋肉痛に注意）について説明グレープフルーツジュース（CYP3A4阻害）回避を指導コンプライアンス向上のため一包化を検討　一包化指導中";
        };

        pasteButton.Click += (s, e) => PasteToWindow();

        Controls.Add(settingsButton);
        Controls.Add(insertVoiceText);
        Controls.Add(lblVoiceText);
        Controls.Add(voiceTextBox);
        Controls.Add(colorDropdown);
        Controls.Add(aiSummarize);
        Controls.Add(lblS);
        Controls.Add(subjectiveBox);
        Controls.Add(lblO);
        Controls.Add(objectiveBox);
        Controls.Add(lblA);
        Controls.Add(assessmentBox);
        Controls.Add(lblP);
        Controls.Add(planBox);
        Controls.Add(pasteButton);
    }

    private void PasteToWindow()
    {
        string[] texts =
        [
            subjectiveBox.Text,
            objectiveBox.Text,
            assessmentBox.Text,
            planBox.Text
        ];

        // 少なくとも1つのフィールドにコンテンツがあるかどうかを確認
        bool hasContent = false;
        foreach (string text in texts)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                hasContent = true;
                break;
            }
        }

        if (!hasContent)
        {
            MessageBox.Show("少なくとも1つのフィールドに入力してください。", "入力が必要です",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 貼付対象ウィンドウを探す
        IntPtr hWnd = WindowHelpers.FindWindowByClassAndTitle(WindowClass, WindowTitle);
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
        Thread.Sleep(300); //  貼付対象ウィンドウがフォーカスされるまで待ち

        // 各フィールドを200ミリ秒の遅延で貼り付けます
        for (int i = 0; i < texts.Length; i++)
        {

            // フィールドが空の場合はスペースを使用
            string textToPaste = string.IsNullOrWhiteSpace(texts[i]) ? " " : texts[i];

            // STAスレッドでクリップボードを設定
            var t = new Thread(() => Clipboard.SetText(textToPaste));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            Thread.Sleep(200); // 貼り付け前に200msの遅延

            // 貼り付け
            SendKeys.SendWait("^v");

            // 次のフィールドに移動する（最後のフィールドの後を除く）
            if (i < texts.Length - 1)
            {
                int pressCountInt = int.Parse(PressCount);
                for (int j = 0; j < pressCountInt; j++)
                {
                    int delayInt = int.Parse(Delay);
                    Thread.Sleep(delayInt); // 遷移の前に遅延
                    switch (ButtonType?.Trim().ToLower())
                    {
                        case "tab":
                            SendKeys.SendWait("{TAB}");
                            break;
                        case "shift":
                            SendKeys.SendWait("+");
                            break;
                        case "ctrl":
                            SendKeys.SendWait("^");
                            break;
                        case "alt":
                            SendKeys.SendWait("%");
                            break;
                        default:
                            MessageBox.Show("設定で項目遷移を設定してください");
                            break;
                    }
                }
            }
        }
    }

    private void PasteToWindowFormm_Load(object sender, EventArgs e)
    {
        var settings = AppSettings.Get();
        WindowTitle = settings.CurrentWindowTitle;
        WindowClass = settings.CurrentWindowClass;
        ButtonType = settings.CurrentButtonType;
        PressCount = settings.CurrentPressCount;
        Delay = settings.CurrentDelay;
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.Run(new PasteToWindowForm());
    }
}
