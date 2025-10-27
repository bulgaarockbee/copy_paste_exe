using System;
using System.Drawing;
using System.Windows.Forms;

public class SettingsForm : Form
{
    public SettingsForm()
    {
        this.Text = "設定";
        this.Size = new Size(300, 200);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;

        Label label = new Label()
        {
            Text = "貼付対象設定:",
            Left = 20,
            Top = 20,
            Width = 100
        };

        Label lblWindowName = new Label() { Text = "貼付対象", Left = 10, Top = 710, Width = 70 };

        Label lblTitle = new Label() { Text = "Title:", Left = 90, Top = 710, Width = 40 };

        titleBox = new TextBox()
        {
            Left = 130,
            Top = 710,
            Width = 200,
            Height = 20,
            Multiline = true,
        };

        Label lblClass = new Label() { Text = "Class:", Left = 340, Top = 710, Width = 40 };

        classBox = new TextBox()
        {
            Left = 380,
            Top = 710,
            Width = 200,
            Height = 20,
            Multiline = true,
        };

        windowCheckButton = new Button()
        {
            Text = "チェック",
            Left = 600,
            Top = 710,
            Width = 90,
            Height = 20,
            BackColor = Color.CornflowerBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        windowCheckButton.FlatAppearance.BorderColor = Color.Black;
        windowCheckButton.FlatAppearance.BorderSize = 1;
        
        Button saveButton = new Button()
        {
            Text = "保存",
            Left = 180,
            Top = 120,
            Width = 80,
            BackColor = Color.LightSkyBlue,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        saveButton.FlatAppearance.BorderSize = 0;

        saveButton.Click += (sender, e) =>
        {
            MessageBox.Show("Settings saved!");
            this.Close();
        };

        this.Controls.Add(label);
        this.Controls.Add(darkModeCheckbox);
        this.Controls.Add(saveButton);
        this.Controls.Add(lblTitle);
        this.Controls.Add(titleBox);
        this.Controls.Add(lblClass);
        this.Controls.Add(classBox);
        this.Controls.Add(windowCheckButton);
    }
}
