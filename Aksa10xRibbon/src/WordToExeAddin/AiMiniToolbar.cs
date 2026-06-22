using System;
using System.Drawing;
using System.Windows.Forms;

public class AiMiniToolbar : Form
{
    public event Action<string> CommandChosen;

    public AiMiniToolbar()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        int w = 214;
        Size = new Size(w, 34);
        MinimumSize = Size;
        MaximumSize = Size;
        BackColor = Color.FromArgb(50, 50, 50);

        int x = 4;
        foreach (var cmd in new[] { "Write", "Fix", "Summarize" })
        {
            var btn = new Button
            {
                Text = cmd,
                Location = new Point(x, 4),
                Size = new Size(66, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 70),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.MouseEnter += (s, e) => btn.BackColor = Color.FromArgb(0, 100, 200);
            btn.MouseLeave += (s, e) => btn.BackColor = Color.FromArgb(70, 70, 70);
            string captured = cmd;
            btn.Click += (s, e) => { if (CommandChosen != null) CommandChosen(captured); };
            Controls.Add(btn);
            x += 70;
        }
    }
}
