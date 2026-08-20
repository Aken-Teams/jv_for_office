// JVision for Excel — 安裝程式
//
// 為什麼是一個 .exe 而不是原本的 .cmd
// ------------------------------------
// 「安裝」本身只有三個動作（複製 manifest、寫一個 HKCU 登錄值、登記到「應用程式
// 和功能」），install.ps1 做得好好的。改成 exe 是為了兩件 .cmd 做不到的事：
//
//   1. 企業派送。這個增益集的主要安裝者不是雙擊檔案的人，而是拿它去推給全公司的
//      infra。派送需要的是靜默安裝、命令列參數、離開代碼和記錄檔——用 Intune、
//      SCCM 或登入指令碼推一個會跳出主控台視窗的 .cmd 是不能看的。
//   2. 自己裝的人看到的是一個視窗，不是一閃而過的黑畫面。
//
// 所以同一個 exe 有兩種面貌：不帶參數是安裝畫面，帶 /silent 是無聲執行。
//
// 用 WinForms 而不是 WPF，因為 csc.exe（每台 Windows 都內建的 .NET Framework
// 編譯器）編得動它——整個建置不需要安裝任何工具，infra 拿到原始碼也重建得出來。
//
// manifest.xml 是嵌進 exe 的，所以發出去的就是「一個檔案」。伺服器網址可以用
// /url 覆寫，同一個 exe 因此可以裝到不同環境，不必為每個環境維護一份 manifest。
//
// 全部只寫 HKCU 和 %LOCALAPPDATA%，不需要系統管理員權限。這也表示它必須在
// 「使用者的身分」下執行——以 SYSTEM 派送會裝到 SYSTEM 自己的登錄檔裡，等於沒裝。
//
// 這個 exe 同時是解除安裝程式：/uninstall 就是移除，「應用程式和功能」裡那顆按鈕
// 呼叫的就是它自己。
//
// 建置：tools\build-installer.ps1
// 注意：csc.exe 是 C# 5，不能用字串內插（$"..."）、?. 或運算式主體成員。

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

namespace JVisionSetup
{
    internal static class Program
    {
        public const string AppName = "JVision for Excel";
        public const string ShortName = "JVisionForExcel";
        /// <summary>開發用網址。manifest 裡所有這個字串都會被換成 /url 指定的位址。</summary>
        public const string DevUrl = "https://localhost:3443";

        public const string WefKey = @"SOFTWARE\Microsoft\Office\16.0\Wef\Developer";
        public const string OfficeKey = @"SOFTWARE\Microsoft\Office\16.0";
        public const string UninstallKey =
            @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ShortName;

        public static string InstallDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppName);
            }
        }

        [STAThread]
        private static int Main(string[] args)
        {
            Options o = Options.Parse(args);
            Log.Path = o.LogPath;
            Log.Write("start: " + string.Join(" ", args));

            try
            {
                if (o.Help) return ShowHelp();
                if (o.Uninstall) return RunUninstall(o);
                if (o.Silent) return RunSilentInstall(o);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new SetupForm(o));
                return ExitCode.Ok;
            }
            catch (Exception ex)
            {
                Log.Write("FATAL: " + ex);
                if (!o.Silent)
                    MessageBox.Show(ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ExitCode.Failed;
            }
        }

        private static int ShowHelp()
        {
            MessageBox.Show(
                AppName + " 安裝程式\n\n" +
                "直接執行：開啟安裝畫面\n\n" +
                "命令列（給 IT 派送用）：\n" +
                "  /silent              不顯示畫面，直接安裝\n" +
                "  /url:<網址>          指定任務窗格伺服器\n" +
                "  /log:<檔案路徑>      寫入記錄檔\n" +
                "  /uninstall           移除\n\n" +
                "離開代碼：0 成功、1 失敗、2 這台電腦沒有 Office 16.0、3 使用者取消\n\n" +
                "必須在使用者身分下執行（設定寫在 HKCU）。",
                AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return ExitCode.Ok;
        }

        private static int RunUninstall(Options o)
        {
            if (!o.Silent)
            {
                DialogResult ok = MessageBox.Show(
                    "要移除 " + AppName + " 嗎？\n\n" +
                    "這只會把增益集從 Excel 移除，不會動到任何活頁簿。",
                    AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (ok != DialogResult.Yes) return ExitCode.Cancelled;
            }
            try
            {
                Installer.Uninstall();
                Log.Write("uninstalled");
                if (!o.Silent)
                    MessageBox.Show("已移除。請重新啟動 Excel。",
                        AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return ExitCode.Ok;
            }
            catch (Exception ex)
            {
                Log.Write("uninstall failed: " + ex);
                if (!o.Silent)
                    MessageBox.Show("移除失敗：" + ex.Message,
                        AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ExitCode.Failed;
            }
        }

        private static int RunSilentInstall(Options o)
        {
            // Not an error: a machine without Excel 2016+ simply isn't a target. A
            // deployment can map this code to "not applicable" instead of alerting
            // someone at 3am about a laptop that was never going to run this.
            if (!Installer.HasOffice16())
            {
                Log.Write("Office 16.0 not present for this user - nothing to do");
                return ExitCode.NoOffice;
            }
            try
            {
                Installer.Install(o.Url);
                Log.Write("installed to " + InstallDir + (o.Url == null ? "" : "  url=" + o.Url));
                return ExitCode.Ok;
            }
            catch (Exception ex)
            {
                Log.Write("install failed: " + ex);
                return ExitCode.Failed;
            }
        }
    }

    internal static class ExitCode
    {
        public const int Ok = 0;
        public const int Failed = 1;
        /// <summary>No Office 16.0 for this user — a deployment should treat this as "not applicable".</summary>
        public const int NoOffice = 2;
        public const int Cancelled = 3;
    }

    /// <summary>
    /// 命令列參數。這個 exe 的主要使用者不是雙擊它的人，而是拿它去派送的 infra，
    /// 所以參數是一等公民，不是附加功能。
    /// </summary>
    internal sealed class Options
    {
        public bool Silent;
        public bool Uninstall;
        public bool Help;
        public string Url;
        public string LogPath;

        public static Options Parse(string[] args)
        {
            Options o = new Options();
            foreach (string a in args)
            {
                string s = a.TrimStart('-', '/');
                if (Is(s, "silent") || Is(s, "quiet") || Is(s, "s") || Is(s, "q")) o.Silent = true;
                else if (Is(s, "uninstall") || Is(s, "x")) o.Uninstall = true;
                else if (Is(s, "help") || Is(s, "?") || Is(s, "h")) o.Help = true;
                else if (Starts(s, "url")) o.Url = Value(s);
                else if (Starts(s, "log")) o.LogPath = Value(s);
            }
            return o;
        }

        private static bool Is(string s, string name)
        {
            return s.Equals(name, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Accepts both /url:x and /url=x — IT tooling writes either.</summary>
        private static bool Starts(string s, string name)
        {
            return s.StartsWith(name + ":", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase);
        }

        private static string Value(string s)
        {
            int i = s.IndexOfAny(new[] { ':', '=' });
            return s.Substring(i + 1).Trim().Trim('"');
        }
    }

    /// <summary>
    /// Writes to /log if given, otherwise %TEMP%\JVisionSetup.log.
    ///
    /// A silent install that fails has no screen to report on. This file is the
    /// only place anyone can find out why, so it is written even when /log is not
    /// passed — a default location beats no record at all.
    /// </summary>
    internal static class Log
    {
        public static string Path;

        public static void Write(string line)
        {
            try
            {
                string p = Path;
                if (string.IsNullOrEmpty(p))
                    p = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "JVisionSetup.log");
                File.AppendAllText(p,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + line + Environment.NewLine,
                    new UTF8Encoding(false));
            }
            catch { }
        }
    }

    /// <summary>
    /// 安裝／移除的實際動作。跟畫面完全分開，因為靜默路徑不開視窗。
    /// </summary>
    internal static class Installer
    {
        /// <summary>
        /// 嵌在 exe 裡的圖片資源。找不到就回 null——少一張圖不該讓安裝失敗。
        /// </summary>
        public static Image ReadEmbeddedImage(string suffix)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                string name = asm.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
                if (name == null) return null;
                using (Stream s = asm.GetManifestResourceStream(name))
                {
                    // Image.FromStream keeps a reference to the stream, so copy first.
                    MemoryStream ms = new MemoryStream();
                    s.CopyTo(ms);
                    ms.Position = 0;
                    return Image.FromStream(ms);
                }
            }
            catch { return null; }
        }

        /// <summary>嵌在 exe 裡的 manifest 原文。</summary>
        public static string ReadEmbeddedManifest()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string name = asm.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith("manifest.xml", StringComparison.OrdinalIgnoreCase));
            if (name == null)
                throw new InvalidOperationException("這個安裝程式沒有內嵌 manifest.xml，請重新建置。");
            using (Stream s = asm.GetManifestResourceStream(name))
            using (StreamReader r = new StreamReader(s, new UTF8Encoding(false)))
            {
                return r.ReadToEnd();
            }
        }

        /// <summary>
        /// 取 manifest 的一個頂層欄位。刻意忽略命名空間——manifest 的 xmlns 換過幾次，
        /// 綁死命名空間只會在下一次改版時安靜地讀不到值。
        /// </summary>
        public static string Field(XDocument doc, string localName)
        {
            XElement el = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == localName);
            return el == null ? null : el.Value.Trim();
        }

        /// <summary>manifest 目前指向的伺服器（只要 scheme + host，不要後面的路徑）。</summary>
        public static string TaskpaneOrigin(XDocument doc)
        {
            XElement ds = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "DefaultSettings");
            if (ds == null) return null;
            XElement sl = ds.Elements().FirstOrDefault(e => e.Name.LocalName == "SourceLocation");
            if (sl == null) return null;
            XAttribute v = sl.Attributes().FirstOrDefault(a => a.Name.LocalName == "DefaultValue");
            if (v == null) return null;
            try { return new Uri(v.Value).GetLeftPart(UriPartial.Authority); }
            catch { return v.Value; }
        }

        public static bool HasOffice16()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(Program.OfficeKey)) { return k != null; }
        }

        public static bool IsInstalled()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(Program.UninstallKey)) { return k != null; }
        }

        public static bool ExcelIsRunning()
        {
            try { return Process.GetProcessesByName("EXCEL").Length > 0; }
            catch { return false; }
        }

        /// <summary>安裝。baseUrl 為空就沿用 manifest 原本的網址。</summary>
        public static void Install(string baseUrl)
        {
            string raw = ReadEmbeddedManifest();
            XDocument doc = XDocument.Parse(raw);

            string id = Field(doc, "Id");
            string version = Field(doc, "Version");
            string publisher = Field(doc, "ProviderName");
            if (string.IsNullOrEmpty(id))
                throw new InvalidOperationException("manifest.xml 裡沒有 <Id>。");

            if (!HasOffice16())
                throw new InvalidOperationException(
                    "這台電腦沒有偵測到 Office 16.0（Microsoft 365 / Office 2016 以後）。\n" +
                    Program.AppName + " 需要 Microsoft 365 版的 Excel。");

            Directory.CreateDirectory(Program.InstallDir);

            // 換網址。開發用的 localhost 和先前裝過的內網位址都會被換掉，
            // 所以重跑安裝程式換環境是安全的。
            if (!string.IsNullOrEmpty(baseUrl))
            {
                string trimmed = baseUrl.TrimEnd('/');
                // $ in the replacement is a backreference to Regex.Replace, and an
                // URL can legitimately contain one.
                raw = Regex.Replace(raw, Regex.Escape(Program.DevUrl), trimmed.Replace("$", "$$"));
            }

            // manifest 一定要放在穩定的位置：留在下載資料夾的檔案會被清掉，
            // 之後 Excel 就載入失敗，而且失敗得很安靜。
            string manifestPath = Path.Combine(Program.InstallDir, "manifest.xml");
            File.WriteAllText(manifestPath, raw, new UTF8Encoding(false));

            // 把自己複製進去，「應用程式和功能」那顆解除安裝按鈕才不會在使用者
            // 清掉下載資料夾之後失效。
            string selfTarget = Path.Combine(Program.InstallDir, "JVisionSetup.exe");
            string selfSource = Assembly.GetExecutingAssembly().Location;
            if (!string.Equals(selfSource, selfTarget, StringComparison.OrdinalIgnoreCase))
                File.Copy(selfSource, selfTarget, true);

            // Excel 認的就是這一個值：名稱是增益集 Id，資料是 manifest 的絕對路徑。
            using (RegistryKey wef = Registry.CurrentUser.CreateSubKey(Program.WefKey))
            {
                // 舊版工具用 manifest 路徑當值名稱。兩種都清掉再寫，
                // 免得同一個增益集在 Excel 的清單裡出現兩次。
                try { wef.DeleteValue(manifestPath, false); }
                catch { }
                wef.SetValue(id, manifestPath, RegistryValueKind.String);
            }

            using (RegistryKey u = Registry.CurrentUser.CreateSubKey(Program.UninstallKey))
            {
                u.SetValue("DisplayName", Program.AppName);
                if (!string.IsNullOrEmpty(version)) u.SetValue("DisplayVersion", version);
                if (!string.IsNullOrEmpty(publisher)) u.SetValue("Publisher", publisher);
                u.SetValue("InstallLocation", Program.InstallDir);
                u.SetValue("DisplayIcon", selfTarget);
                u.SetValue("UninstallString", "\"" + selfTarget + "\" /uninstall");
                u.SetValue("QuietUninstallString", "\"" + selfTarget + "\" /uninstall /silent");
                u.SetValue("NoModify", 1, RegistryValueKind.DWord);
                u.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
        }

        public static void Uninstall()
        {
            string manifestPath = Path.Combine(Program.InstallDir, "manifest.xml");

            string id = null;
            try
            {
                if (File.Exists(manifestPath)) id = Field(XDocument.Load(manifestPath), "Id");
            }
            catch { }

            using (RegistryKey wef = Registry.CurrentUser.OpenSubKey(Program.WefKey, true))
            {
                if (wef != null)
                {
                    if (!string.IsNullOrEmpty(id)) { try { wef.DeleteValue(id, false); } catch { } }
                    // Belt and braces: if the manifest was already gone we cannot look
                    // the Id up, and a value left pointing at a deleted file makes
                    // Excel fail to load an add-in that no longer exists.
                    foreach (string name in wef.GetValueNames())
                    {
                        object val = wef.GetValue(name);
                        string s = val == null ? "" : val.ToString();
                        if (s.StartsWith(Program.InstallDir, StringComparison.OrdinalIgnoreCase))
                        {
                            try { wef.DeleteValue(name, false); }
                            catch { }
                        }
                    }
                }
            }

            try { Registry.CurrentUser.DeleteSubKeyTree(Program.UninstallKey, false); }
            catch { }

            // We are running from inside the folder we are deleting, so hand the last
            // step to a detached cmd that waits for this process to exit.
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe",
                    "/c timeout /t 2 /nobreak >nul & rd /s /q \"" + Program.InstallDir + "\"");
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;
                Process.Start(psi);
            }
            catch { }
        }
    }

    /// <summary>
    /// 安裝畫面。給自己裝的人看的那一面。
    /// </summary>
    internal sealed class SetupForm : Form
    {
        private readonly Options _opts;
        private TextBox _url;
        private Button _install;
        private Button _remove;
        private Label _status;

        public SetupForm(Options opts)
        {
            _opts = opts;
            BuildUi();
        }

        private void BuildUi()
        {
            Text = Program.AppName;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(430, 250);
            Font = new Font("Microsoft JhengHei UI", 9F);
            BackColor = Color.White;

            string version = "";
            string origin = "";
            try
            {
                XDocument doc = XDocument.Parse(Installer.ReadEmbeddedManifest());
                version = Installer.Field(doc, "Version");
                origin = Installer.TaskpaneOrigin(doc);
            }
            catch { }

            // The mark, at the size Explorer and the Store use. An installer with no
            // branding is indistinguishable from something that arrived by accident,
            // which is the whole objection this exe exists to answer.
            Image logo = Installer.ReadEmbeddedImage("logo.png");
            int textLeft = 24;
            if (logo != null)
            {
                PictureBox pb = new PictureBox();
                pb.Image = logo;
                pb.SizeMode = PictureBoxSizeMode.Zoom;
                pb.SetBounds(24, 20, 48, 48);
                Controls.Add(pb);
                textLeft = 84;
            }

            Label title = new Label();
            title.Text = Program.AppName;
            title.Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold);
            title.ForeColor = Color.FromArgb(16, 124, 65);
            title.SetBounds(textLeft, 22, 330, 32);
            Controls.Add(title);

            Label sub = new Label();
            sub.Text = string.IsNullOrEmpty(version)
                ? "Excel 增益集"
                : "Excel 增益集　版本 " + version;
            sub.ForeColor = Color.FromArgb(96, 96, 96);
            sub.SetBounds(textLeft + 2, 54, 330, 20);
            Controls.Add(sub);

            Label urlLabel = new Label();
            urlLabel.Text = "伺服器網址";
            urlLabel.SetBounds(26, 92, 100, 20);
            Controls.Add(urlLabel);

            _url = new TextBox();
            // Pre-filled from the manifest so the common case is "just press install",
            // editable so one exe serves every environment.
            _url.Text = _opts.Url != null ? _opts.Url : (origin != null ? origin : "");
            _url.SetBounds(26, 112, 378, 26);
            Controls.Add(_url);

            _status = new Label();
            _status.SetBounds(26, 148, 378, 40);
            _status.ForeColor = Color.FromArgb(96, 96, 96);
            Controls.Add(_status);

            _install = new Button();
            _install.Text = Installer.IsInstalled() ? "重新安裝" : "安裝";
            _install.SetBounds(284, 198, 120, 34);
            _install.BackColor = Color.FromArgb(16, 124, 65);
            _install.ForeColor = Color.White;
            _install.FlatStyle = FlatStyle.Flat;
            _install.FlatAppearance.BorderSize = 0;
            _install.Click += OnInstall;
            Controls.Add(_install);
            AcceptButton = _install;

            _remove = new Button();
            _remove.Text = "移除";
            _remove.SetBounds(156, 198, 116, 34);
            _remove.FlatStyle = FlatStyle.Flat;
            _remove.Visible = Installer.IsInstalled();
            _remove.Click += OnRemove;
            Controls.Add(_remove);

            if (!Installer.HasOffice16())
            {
                _install.Enabled = false;
                Say("這台電腦沒有偵測到 Microsoft 365 版的 Excel。", true);
            }
        }

        private void Say(string text, bool warn)
        {
            _status.ForeColor = warn ? Color.FromArgb(164, 38, 44) : Color.FromArgb(96, 96, 96);
            _status.Text = text;
            _status.Refresh();
        }

        private void OnInstall(object sender, EventArgs e)
        {
            _install.Enabled = false;
            Say("安裝中…", false);
            try
            {
                Installer.Install(_url.Text.Trim());
                Log.Write("installed via UI");
                // Excel only reads the registry at startup, so an open Excel will not
                // show the button. Deliberately not killing it — there could be
                // unsaved work in there.
                string msg = Installer.ExcelIsRunning()
                    ? "安裝完成。\n\nExcel 目前是開著的，請存檔後完全關閉 Excel 再重新開啟，「常用」分頁最右邊就會出現 JVision。"
                    : "安裝完成。\n\n開啟 Excel 後，在「常用」分頁最右邊會看到 JVision。";
                MessageBox.Show(msg, Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            catch (Exception ex)
            {
                Log.Write("install failed via UI: " + ex);
                Say(ex.Message, true);
                _install.Enabled = true;
            }
        }

        private void OnRemove(object sender, EventArgs e)
        {
            DialogResult ok = MessageBox.Show(
                "要移除 " + Program.AppName + " 嗎？\n\n" +
                "這只會把增益集從 Excel 移除，不會動到任何活頁簿。",
                Program.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ok != DialogResult.Yes) return;
            try
            {
                Installer.Uninstall();
                MessageBox.Show("已移除。請重新啟動 Excel。",
                    Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            catch (Exception ex)
            {
                Say("移除失敗：" + ex.Message, true);
            }
        }
    }
}
