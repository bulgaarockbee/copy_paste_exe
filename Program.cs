using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

class PasteToChromeForm : Form
{
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

    [DllImport("user32.dll")]
    private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder text, int count);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    const int SW_RESTORE = 9;

    private TextBox subjectiveBox;
    private TextBox objectiveBox;
    private TextBox assessmentBox;
    private TextBox planBox;
    private TextBox titleBox;
    private TextBox classBox;
    private TextBox voiceTextBox;
    private Button pasteButton;
    private Button windowCheckButton;
    private Button insertVoiceText;
    private Button aiSummarize;

    public PasteToChromeForm()
    {
        this.Text = "S.O.A.Pコピー";
        this.Width = 720;
        this.Height = 780;

        Button settingsButton = new Button()
        {
            Text = "⚙️",
            Left = 650,
            Top = 10,
            Width = 30,
            Height = 30,
            BackColor = Color.LightGray,
            FlatStyle = FlatStyle.Flat
        };

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

        Label lblVoiceText = new Label() { Text = "音声書き起こし:", Left = 10, Top = 50, Width = 300 };
        voiceTextBox = new TextBox()
        {
            Left = 10,
            Top = 75,
            Width = 680,
            Height = 100,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        ComboBox colorDropdown = new ComboBox()
        {
            Left = 200,
            Top = 185,
            Width = 150,
            DropDownStyle = ComboBoxStyle.DropDownList
        };

        colorDropdown.Items.AddRange(new string[]
        {
            "SOAP(アシスト)"
        });

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

        Label lblS = new Label() { Text = "主観的情報（Subjective）:", Left = 10, Top = 210, Width = 300 };
        subjectiveBox = new TextBox()
        {
            Left = 10,
            Top = 235,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblO = new Label() { Text = "客観的情報（Objective）:", Left = 10, Top = 320, Width = 300 };
        objectiveBox = new TextBox()
        {
            Left = 10,
            Top = 345,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblA = new Label() { Text = "評価（Assessment）:", Left = 10, Top = 430, Width = 300 };
        assessmentBox = new TextBox()
        {
            Left = 10,
            Top = 455,
            Width = 680,
            Height = 80,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        Label lblP = new Label() { Text = "計画（Plan）:", Left = 10, Top = 540, Width = 300 };
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
            SettingsForm settingsForm = new SettingsForm();
            settingsForm.ShowDialog();  // Show as modal popup
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

        pasteButton.Click += (s, e) => PasteToChrome();

        this.Controls.Add(settingsButton);
        this.Controls.Add(insertVoiceText);
        this.Controls.Add(lblVoiceText);
        this.Controls.Add(voiceTextBox);
        this.Controls.Add(colorDropdown);
        this.Controls.Add(aiSummarize);
        this.Controls.Add(lblS);
        this.Controls.Add(subjectiveBox);
        this.Controls.Add(lblO);
        this.Controls.Add(objectiveBox);
        this.Controls.Add(lblA);
        this.Controls.Add(assessmentBox);
        this.Controls.Add(lblP);
        this.Controls.Add(planBox);
        this.Controls.Add(pasteButton);
    }

    private void CheckWindow()
    {
        IntPtr hWnd = FindWindowByClassAndTitle();
        if (hWnd == IntPtr.Zero)
        {
            MessageBox.Show("貼付対象ウィンドウが見つかりません。まずウィンドウを開くまたはタイトルやクラス名を直してください。", "ウィンドウが見つかりません",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Chromeを復元してフォーカスする
        if (IsIconic(hWnd))
        {
            ShowWindow(hWnd, SW_RESTORE);
        }
        SetForegroundWindow(hWnd);
    }

    private void PasteToChrome()
    {
        string[] texts = new string[]
        {
            subjectiveBox.Text,
            objectiveBox.Text,
            assessmentBox.Text,
            planBox.Text
        };

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

        // Chrome ウィンドウを探す
        // IntPtr hWnd = FindChromeWindow();
        IntPtr hWnd = FindWindowByClassAndTitle();
        if (hWnd == IntPtr.Zero)
        {
            MessageBox.Show("貼付対象ウィンドウが見つかりません。まずウィンドウを開くまたはタイトルやクラス名を直してください。", "ウィンドウが見つかりません",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Chromeを復元してフォーカスする
        if (IsIconic(hWnd))
        {
            ShowWindow(hWnd, SW_RESTORE);
        }
        SetForegroundWindow(hWnd);
        Thread.Sleep(300); // Chromeがフォーカスされるまで待ち

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
                Thread.Sleep(200); // タブの前に200msの遅延
                SendKeys.SendWait("{TAB}");
            }
        }
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.Run(new PasteToChromeForm());
    }
}
