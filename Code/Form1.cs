using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Xml;

namespace XDocUnlocker
{
    public partial class Form1 : Form
    {
        private string selectedFilePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbBackup.Checked = true;
        }

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Office (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx";
                dlg.Title = "Choose Locked Office-File";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = dlg.FileName;
                    tbPath.Text = selectedFilePath;  // VOLLER PFAD
                    tbPath.SelectionStart = tbPath.Text.Length;  // Cursor ans Ende
                    tbPath.SelectionLength = 0;
                    tbPath.ScrollToCaret();

                    BtnUnlock.Enabled = true;
                    lblStatus.Text = "Bereit.";
                    lblStatus.ForeColor = System.Drawing.Color.Green;
                }
            }
        }

        // TextChanged Event für tbPath
        private void tbPath_TextChanged(object sender, EventArgs e)
        {
            selectedFilePath = tbPath.Text;
        }

        private void BtnUnlock_Click(object sender, EventArgs e)
        {
            selectedFilePath = tbPath.Text;
            // Datei existiert?
            if (!File.Exists(selectedFilePath))
            {
                MessageBox.Show("Datei existiert nicht!\nBitte gültigen Pfad eingeben.", "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                BtnUnlock.Enabled = false;
                lblStatus.Text = "Entferne Schutz...";
                lblStatus.ForeColor = System.Drawing.Color.Orange;
                this.Refresh();

                string unlockedPath = Path.Combine(
                    Path.GetDirectoryName(selectedFilePath),
                    Path.GetFileNameWithoutExtension(selectedFilePath) + "_unlocked" +
                    Path.GetExtension(selectedFilePath));

                string backupPath = "";
                if (cbBackup.Checked)
                {
                    backupPath = selectedFilePath + ".backup";
                    lblStatus.Text = "Erstelle Backup...";
                    this.Refresh();
                    File.Copy(selectedFilePath, backupPath, true);
                }

                RemoveProtection(selectedFilePath, unlockedPath);

                string successMsg = $"Schutz erfolgreich entfernt!\nEntsperrte Datei: {Path.GetFileName(unlockedPath)}";
                if (cbBackup.Checked)
                    successMsg += $"\nBackup: {Path.GetFileName(backupPath)}";

                lblStatus.Text = "Erfolg!";
                lblStatus.ForeColor = System.Drawing.Color.Green;
                MessageBox.Show(successMsg, "XDocUnlocker", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Fehler: " + ex.Message;
                lblStatus.ForeColor = System.Drawing.Color.Red;
                MessageBox.Show("Fehler beim Entsperren:\n" + ex.Message, "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BtnUnlock.Enabled = true;
            }
        }

        // Rest des Codes unverändert...
        private void RemoveProtection(string inputPath, string outputPath)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "XDocUnlocker_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(tempDir);

            try
            {
                ZipFile.ExtractToDirectory(inputPath, tempDir);
                string ext = Path.GetExtension(inputPath).ToLower();

                switch (ext)
                {
                    case ".xlsx": ProcessExcel(tempDir); break;
                    case ".docx": ProcessWord(tempDir); break;
                    case ".pptx": ProcessPowerPoint(tempDir); break;
                }

                File.Copy(inputPath, outputPath, true);
                using (var zip = ZipFile.Open(outputPath, ZipArchiveMode.Update))
                {
                    PatchZipEntries(tempDir, zip, ext);
                }
            }
            finally
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, true);
            }
        }

        private void ProcessExcel(string tempDir)
        {
            string xlDir = Path.Combine(tempDir, "xl", "worksheets");
            if (Directory.Exists(xlDir))
            {
                foreach (string sheet in Directory.GetFiles(xlDir, "sheet*.xml"))
                    RemoveProtectionTag(sheet, "sheetProtection");
            }
        }

        private void ProcessWord(string tempDir)
        {
            string wordSettings = Path.Combine(tempDir, "word", "settings.xml");
            if (File.Exists(wordSettings))
                RemoveProtectionTag(wordSettings, "w:documentProtection");
        }

        private void ProcessPowerPoint(string tempDir)
        {
            string pptFile = Path.Combine(tempDir, "ppt", "presentation.xml");
            if (File.Exists(pptFile))
                RemoveProtectionTag(pptFile, "p:modifyVerifier");
        }

        private void PatchZipEntries(string tempDir, ZipArchive zip, string ext)
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

        private void RemoveProtectionTag(string xmlFile, string tagName)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.PreserveWhitespace = true;
                doc.Load(xmlFile);

                XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                if (tagName == "sheetProtection")
                {
                    XmlNodeList sheets = doc.SelectNodes("//x:sheetProtection", ns);
                    foreach (XmlNode node in sheets)
                        node.ParentNode?.RemoveChild(node);
                }
                else if (tagName == "w:documentProtection")
                {
                    XmlNodeList prot = doc.SelectNodes("//w:documentProtection", ns);
                    foreach (XmlNode node in prot)
                    {
                        if (node.Attributes["w:enforcement"] is XmlAttribute enforce)
                            enforce.Value = "0";
                    }
                }
                else if (tagName == "p:modifyVerifier")
                {
                    XmlNodeList mods = doc.SelectNodes("//p:modifyVerifier", ns);
                    foreach (XmlNode node in mods)
                        node.ParentNode?.RemoveChild(node);
                }

                doc.Save(xmlFile);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fehler beim Verarbeiten {xmlFile}: {ex.Message}");
            }
        }
    }
}
