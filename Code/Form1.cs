using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Drawing;
using System.Collections.Generic;

namespace XDocUnlocker
{
    public partial class Form1 : Form
    {
        private string? selectedFilePath = "";
        private Point _mouseDownPos;
        private bool _isPinned = false;

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbBackup.Checked = true;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = 100;
            progressBar1.Visible = false;

            lblStatus.Text = "💡 Drag & drop Office files or use the button";
            lblStatus.ForeColor = System.Drawing.Color.Blue;
        }

        private void Form1_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
                {
                    bool isOfficeFile = files.Any(f =>
                        f?.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) == true ||
                        f?.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) == true ||
                        f?.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) == true);

                    e.Effect = isOfficeFile ? DragDropEffects.Copy : DragDropEffects.None;
                }
            }
        }

        private void Form1_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            {
                var officeFiles = files
                    .Where(f => !string.IsNullOrEmpty(f) &&
                        (f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                         f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                         f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)))
                    .ToArray();

                if (officeFiles.Length > 0)
                {
                    selectedFilePath = officeFiles[0];
                    tbPath.Text = string.Join(" | ", officeFiles);
                    tbPath.SelectionStart = tbPath.Text.Length;
                    tbPath.SelectionLength = 0;
                    tbPath.ScrollToCaret();

                    BtnUnlock.Enabled = true;
                    lblStatus.Text = $"✅ {officeFiles.Length} Office file(s) loaded. First: '{Path.GetFileName(officeFiles[0])}'";
                    if (officeFiles.Length > 1)
                        lblStatus.Text += $" ({officeFiles.Length - 1} more ready)";
                    lblStatus.ForeColor = System.Drawing.Color.Green;
                }
                else
                {
                    lblStatus.Text = "❌ No valid Office files (.docx/.xlsx/.pptx) found!";
                    lblStatus.ForeColor = System.Drawing.Color.Orange;
                }
            }
        }

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Office (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx";
                dlg.Multiselect = true;
                dlg.Title = "Choose Locked Office File(s)";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    var officeFiles = dlg.FileNames
                        .Where(f => !string.IsNullOrEmpty(f) &&
                            (f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                             f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                             f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)))
                        .ToArray();

                    if (officeFiles.Length > 0)
                    {
                        selectedFilePath = officeFiles[0];
                        tbPath.Text = string.Join(" | ", officeFiles);
                        tbPath.SelectionStart = tbPath.Text.Length;
                        tbPath.SelectionLength = 0;
                        tbPath.ScrollToCaret();

                        BtnUnlock.Enabled = true;
                        lblStatus.Text = $"✅ {officeFiles.Length} Office file(s) loaded. First: '{Path.GetFileName(officeFiles[0])}'";
                        if (officeFiles.Length > 1)
                            lblStatus.Text += $" ({officeFiles.Length - 1} more ready)";
                        lblStatus.ForeColor = System.Drawing.Color.Green;
                    }
                }
            }
        }

        private void tbPath_TextChanged(object sender, EventArgs e)
        {
            selectedFilePath = tbPath.Text;
        }

        private async void BtnUnlock_Click(object sender, EventArgs e)
        {
            string[] allInputFiles = tbPath.Text.Split(new char[] { '|', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(f => f.Trim())
                .Where(f => !string.IsNullOrEmpty(f) && File.Exists(f))
                .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                           f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                           f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (allInputFiles.Length == 0)
            {
                lblStatus.Text = "❌ No valid Office files found!";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                return;
            }

            BtnUnlock.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = allInputFiles.Length * 100;
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            lblStatus.Text = $"🚀 Unlocking {allInputFiles.Length} file(s)...";
            lblStatus.ForeColor = System.Drawing.Color.Orange;
            await Task.Delay(500);
            Application.DoEvents();

            try
            {
                int totalProgress = 0;
                int successCount = 0;

                foreach (string filePath in allInputFiles)
                {
                    try
                    {
                        lblStatus.Text = $"🔓 [{successCount + 1}/{allInputFiles.Length}] {Path.GetFileName(filePath)}...";
                        Application.DoEvents();

                        string unlockedPath = Path.Combine(
                            Path.GetDirectoryName(filePath) ?? "",
                            Path.GetFileNameWithoutExtension(filePath) + "_unlocked" +
                            Path.GetExtension(filePath));

                        string? backupPath = null;
                        if (cbBackup.Checked)
                        {
                            lblStatus.Text = $"📦 Backup: {Path.GetFileName(filePath)}...";
                            backupPath = filePath + ".backup";
                            await Task.Run(() => File.Copy(filePath, backupPath!, true));
                        }

                        var singleProgress = (IProgress<int>?)null;
                        await RemoveProtectionWithProgress(filePath, unlockedPath, singleProgress);

                        successCount++;
                        totalProgress += 100;
                        progressBar1.Value = totalProgress;

                        lblStatus.Text = $"✅ [{successCount}/{allInputFiles.Length}] {Path.GetFileName(filePath)} done";
                        lblStatus.ForeColor = System.Drawing.Color.Green;
                        await Task.Delay(200);
                        Application.DoEvents();
                    }
                    catch (Exception fileEx)
                    {
                        lblStatus.Text = $"⚠️ [{successCount + 1}/{allInputFiles.Length}] ERROR {Path.GetFileName(filePath)}: {fileEx.Message}";
                        lblStatus.ForeColor = System.Drawing.Color.Orange;
                        await Task.Delay(1000);
                        Application.DoEvents();
                    }
                }

                lblStatus.Text = $"✅ COMPLETE! {successCount}/{allInputFiles.Length} file(s) successfully unlocked!";
                lblStatus.ForeColor = System.Drawing.Color.Green;
                await Task.Delay(5000);
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"❌ CRITICAL ERROR: {ex.Message}";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                await Task.Delay(3000);
            }
            finally
            {
                progressBar1.Visible = false;
                BtnUnlock.Enabled = true;
                Application.DoEvents();
            }
        }

        private async Task RemoveProtectionWithProgress(string inputPath, string outputPath, IProgress<int>? progress)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "XDocUnlocker_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(tempDir);

            try
            {
                progress?.Report(10);
                ZipFile.ExtractToDirectory(inputPath, tempDir);

                string ext = Path.GetExtension(inputPath).ToLowerInvariant();
                int totalSteps = GetFileCount(tempDir, ext);
                int stepSize = 80 / Math.Max(1, totalSteps);

                switch (ext)
                {
                    case ".xlsx": await ProcessExcelAsync(tempDir, progress, stepSize); break;
                    case ".docx": await ProcessWordAsync(tempDir, progress); break;
                    case ".pptx": await ProcessPowerPointAsync(tempDir, progress); break;
                }

                progress?.Report(95);
                File.Copy(inputPath, outputPath, true);
                using (var zip = ZipFile.Open(outputPath, ZipArchiveMode.Update))
                {
                    PatchZipEntries(tempDir, zip, ext);
                }

                progress?.Report(100);
            }
            finally
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, true);
            }
        }

        private static int GetFileCount(string tempDir, string ext)
        {
            return ext switch
            {
                ".xlsx" => Directory.GetFiles(Path.Combine(tempDir, "xl", "worksheets"), "sheet*.xml").Length,
                ".docx" => 1,
                ".pptx" => 1,
                _ => 1
            };
        }

        private async Task ProcessExcelAsync(string tempDir, IProgress<int>? progress, int stepSize)
        {
            string xlDir = Path.Combine(tempDir, "xl", "worksheets");
            if (Directory.Exists(xlDir))
            {
                var sheets = Directory.GetFiles(xlDir, "sheet*.xml");
                int current = 10;
                foreach (string sheet in sheets)
                {
                    RemoveProtectionTag(sheet, "sheetProtection");
                    current += stepSize;
                    progress?.Report(Math.Min(90, current));
                    await Task.Delay(10);
                }
            }
        }

        private async Task ProcessWordAsync(string tempDir, IProgress<int>? progress)
        {
            string wordSettings = Path.Combine(tempDir, "word", "settings.xml");
            if (File.Exists(wordSettings))
            {
                RemoveProtectionTag(wordSettings, "w:documentProtection");
                progress?.Report(90);
            }
            await Task.Delay(10);
        }

        private async Task ProcessPowerPointAsync(string tempDir, IProgress<int>? progress)
        {
            string pptFile = Path.Combine(tempDir, "ppt", "presentation.xml");
            if (File.Exists(pptFile))
            {
                RemoveProtectionTag(pptFile, "p:modifyVerifier");
                progress?.Report(90);
            }
            await Task.Delay(10);
        }

        private static void PatchZipEntries(string tempDir, ZipArchive zip, string ext)
        {
            switch (ext)
            {
                case ".xlsx":
                    string xlDir = Path.Combine(tempDir, "xl", "worksheets");
                    if (Directory.Exists(xlDir))
                    {
                        foreach (string sheet in Directory.GetFiles(xlDir, "sheet*.xml"))
                        {
                            string entryName = "xl/worksheets/" + Path.GetFileName(sheet);
                            var entry = zip.GetEntry(entryName);
                            entry?.Delete();
                            zip.CreateEntryFromFile(sheet, entryName, CompressionLevel.Optimal);
                        }
                    }
                    break;
                case ".docx":
                    string wordSettings = Path.Combine(tempDir, "word", "settings.xml");
                    if (File.Exists(wordSettings))
                    {
                        string entryName = "word/settings.xml";
                        var entry = zip.GetEntry(entryName);
                        entry?.Delete();
                        zip.CreateEntryFromFile(wordSettings, entryName, CompressionLevel.Optimal);
                    }
                    break;
                case ".pptx":
                    string pptFile = Path.Combine(tempDir, "ppt", "presentation.xml");
                    if (File.Exists(pptFile))
                    {
                        string entryName = "ppt/presentation.xml";
                        var entry = zip.GetEntry(entryName);
                        entry?.Delete();
                        zip.CreateEntryFromFile(pptFile, entryName, CompressionLevel.Optimal);
                    }
                    break;
            }
        }

        private static void RemoveProtectionTag(string xmlFile, string tagName)
        {
            try
            {
                var doc = new XmlDocument();
                doc.PreserveWhitespace = true;
                doc.Load(xmlFile);

                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                if (tagName == "sheetProtection")
                {
                    var sheets = doc.SelectNodes("//x:sheetProtection", ns);
                    if (sheets != null)
                    {
                        foreach (XmlNode node in sheets)
                            node.ParentNode?.RemoveChild(node);
                    }
                }
                else if (tagName == "w:documentProtection")
                {
                    var prot = doc.SelectNodes("//w:documentProtection", ns);
                    if (prot != null)
                    {
                        foreach (XmlNode node in prot)
                        {
                            if (node.Attributes?["w:enforcement"] is XmlAttribute enforce)
                                enforce.Value = "0";
                        }
                    }
                }
                else if (tagName == "p:modifyVerifier")
                {
                    var mods = doc.SelectNodes("//p:modifyVerifier", ns);
                    if (mods != null)
                    {
                        foreach (XmlNode node in mods)
                            node.ParentNode?.RemoveChild(node);
                    }
                }

                doc.Save(xmlFile);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing {xmlFile}: {ex.Message}");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string url = "https://github.com/TueftelTyp/XDocUnlocker";
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            if (_isPinned) return;

            if (e.Button == MouseButtons.Left)
            {
                _mouseDownPos = new Point(e.X, e.Y);
            }
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isPinned) return;

            if (e.Button == MouseButtons.Left)
            {
                int deltaX = e.X - _mouseDownPos.X;
                int deltaY = e.Y - _mouseDownPos.Y;
                this.Location = new Point(Location.X + deltaX, Location.Y + deltaY);
            }
        }

        private void btnMini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnPin_Click(object sender, EventArgs e)
        {
            _isPinned = !_isPinned;
            this.TopMost = _isPinned;

            if (_isPinned)
            {
                btnPin.Image = Properties.Resources.unpinned;
            }
            else
            {
                btnPin.Image = Properties.Resources.pinned;
            }
        }
    }
}
