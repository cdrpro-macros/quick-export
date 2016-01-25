using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using Corel.Interop.VGCore;

namespace QuickExport
{
    public partial class wSettings : System.Windows.Window
    {

        public wSettings()
        {
            InitializeComponent();
            LoadDefault();
            LoadPresets();
        }



        private void LoadPresets()
        {
            try
            {
                cbPresets.Items.Clear();

                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                if (xDoc.ChildNodes[1].ChildNodes.Count == 0) return;

                foreach (XmlNode n in xDoc.ChildNodes[1].ChildNodes)
                {
                    var cbItem = new ComboBoxItem {Content = n.Attributes["title"].Value, Tag = n.Attributes["pid"].Value};
                    cbPresets.Items.Add(cbItem);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }


        private void LoadDefault()
        {
            var xDoc = new XmlDocument();
            xDoc.Load(Ui.sPath);
            var root = xDoc.ChildNodes[1];
            var fAttribute = root.Attributes["folder"];
            if (fAttribute != null) tbGeneralFolder.Text = root.Attributes["folder"].Value;

            var defaultDpi = 72;
            if (Ui.dApp.Documents.Count != 0) defaultDpi = Ui.dApp.ActiveDocument.Resolution;

            cbFileFormat.Items.Add("Jpeg");
            cbFileFormat.Items.Add("PNG 24");
            cbFileFormat.Text = "Jpeg";

            cbRange.Items.Add("Selection");
            cbRange.Items.Add("Active page");
            cbRange.Items.Add("All pages");
            cbRange.Text = "Selection";

            LoadColorMode(GetDefColorMode(), false);

            cbSubFormat.Items.Add("Standard (4:2:2)");
            cbSubFormat.Items.Add("Optional (4:4:4)");
            cbSubFormat.Text = "Standard (4:2:2)";

            tbQuality.Text = "80";
            tbBlur.Text = "0";
            tbDPI.Text = defaultDpi.ToString(CultureInfo.InvariantCulture);

            bDelPreset.IsEnabled = false;
            btnSave.IsEnabled = false;
        }


        private void LoadColorMode(string colorMode, bool isPng)
        {
            cbColorMode.Items.Clear();
            cbColorMode.Items.Add("RGB Color");
            if (!isPng) cbColorMode.Items.Add("CMYK Color");
            cbColorMode.Items.Add("Grayscale");
            if (isPng && colorMode == "CMYK Color") colorMode = "RGB Color";
            cbColorMode.Text = colorMode;
        }


        private string GetDefColorMode()
        {
            var defaultColorMode = "RGB Color";
            if (Ui.dApp.Documents.Count != 0)
            {
                switch (Ui.dApp.ActiveDocument.ColorContext.BlendingColorModel)
                {
                    case clrColorModel.clrColorModelCMYK:
                        defaultColorMode = "CMYK Color";
                        break;
                    case clrColorModel.clrColorModelGrayscale:
                        defaultColorMode = "Grayscale";
                        break;
                }
            }
            return defaultColorMode;
        }


        private void cbFileFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbFileFormat.SelectedItem.ToString() == "PNG 24")
            {
                LoadColorMode(GetDefColorMode(), true);
                cbTransparency.IsEnabled = true;

                cbInterlaced.IsEnabled = true;
                cbProgressive.IsEnabled = false;
                cbProgressive.IsChecked = false;
                cbOptimize.IsEnabled = false;
                cbOptimize.IsChecked = false;

                tbQuality.IsEnabled = false;
                tbQuality.Text = "0";
                tbBlur.IsEnabled = false;
                tbBlur.Text = "0";
                cbSubFormat.IsEnabled = false;
            }
            else
            {
                LoadColorMode(GetDefColorMode(), false);
                cbTransparency.IsEnabled = false;
                cbTransparency.IsChecked = false;

                cbInterlaced.IsEnabled = false;
                cbInterlaced.IsChecked = false;

                if (!(bool)cbProgressive.IsChecked) cbOptimize.IsEnabled = true;
                if (!(bool)cbOptimize.IsChecked) cbProgressive.IsEnabled = true;

                tbQuality.IsEnabled = true;
                tbBlur.IsEnabled = true;
                cbSubFormat.IsEnabled = true;
            }
        }



        private void cbColorMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbColorMode.Items.Count < 1) return;
            if (cbColorMode.SelectedItem.ToString() == "CMYK Color") cbOverprint.IsEnabled = true;
            else
            {
                cbOverprint.IsChecked = false;
                cbOverprint.IsEnabled = false;
            }
        }





        private bool ValToBool(string val) { return val == "1" ? true : false; }
        private string BoolToVal(bool b) { return b == true ? "1" : "0"; }






        private void cbPresets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (cbPresets.Items.Count < 1) return;

                bDelPreset.IsEnabled = true;
                btnSave.IsEnabled = true;
                
                var selItem = (ComboBoxItem) cbPresets.SelectedItem;

                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var preset = xDoc.SelectSingleNode("//Preset[@pid = \"" + selItem.Tag.ToString() + "\"]");

                if (preset != null)
                {
                    var fileFormat = preset.Attributes["format"].Value;

                    cbFileFormat.Text = fileFormat;
                    cbRange.Text = preset.Attributes["range"].Value;
                    cbColorMode.Text = preset.Attributes["colormode"].Value;
                    cbAntiAliased.IsChecked = ValToBool(preset.Attributes["antialiased"].Value);
                    tbDPI.Text = preset.Attributes["resolution"].Value;

                    switch (fileFormat)
                    {
                        case "Jpeg":
                            cbOverprint.IsChecked = ValToBool(preset.Attributes["overprint"].Value);
                            tbQuality.Text = preset.Attributes["quality"].Value;
                            tbBlur.Text = preset.Attributes["blur"].Value;
                            cbSubFormat.Text = preset.Attributes["subformat"].Value;
                            cbProgressive.IsChecked = ValToBool(preset.Attributes["progressive"].Value);
                            if ((bool)cbProgressive.IsChecked) cbOptimize.IsEnabled = false;
                            cbOptimize.IsChecked = ValToBool(preset.Attributes["optimize"].Value);
                            if ((bool)cbOptimize.IsChecked) cbProgressive.IsEnabled = false;
                            break;
                        case "PNG 24":
                            cbTransparency.IsChecked = ValToBool(preset.Attributes["transparency"].Value);
                            cbInterlaced.IsChecked = ValToBool(preset.Attributes["interlaced"].Value);
                            break;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }





        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void cbProgressive_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)cbProgressive.IsChecked) cbOptimize.IsEnabled = false;
            else cbOptimize.IsEnabled = true;
        }

        private void cbOptimize_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)cbOptimize.IsChecked) cbProgressive.IsEnabled = false;
            else cbProgressive.IsEnabled = true;
        }







        private void bDelPreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (cbPresets.Items.Count < 1) return;
                var selItem = (ComboBoxItem)cbPresets.SelectedItem;

                if (selItem == null)
                {
                    MessageBox.Show("No selected preset.");
                    return;
                }

                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var preset = xDoc.SelectSingleNode("//Preset[@pid = \"" + selItem.Tag.ToString() + "\"]");

                if (preset != null)
                {
                    xDoc.ChildNodes[1].RemoveChild(preset);
                    xDoc.Save(Ui.sPath);
                    LoadPresets();
                    return;
                }

                MessageBox.Show("Item not found!");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }



        private XmlElement CreateNode(XmlDocument xDoc, string pid, string pName)
        {
            var fileFormat = cbFileFormat.Text;

            var preset = xDoc.CreateElement("Preset");
            preset.SetAttribute("pid", pid);
            preset.SetAttribute("title", pName);
            preset.SetAttribute("format", fileFormat);
            preset.SetAttribute("range", cbRange.Text);
            preset.SetAttribute("colormode", cbColorMode.Text);
            preset.SetAttribute("antialiased", BoolToVal((bool)cbAntiAliased.IsChecked));
            preset.SetAttribute("resolution", tbDPI.Text);

            switch (fileFormat)
            {
                case "Jpeg":
                    preset.SetAttribute("overprint", BoolToVal((bool)cbOverprint.IsChecked));
                    preset.SetAttribute("quality", tbQuality.Text);
                    preset.SetAttribute("blur", tbBlur.Text);
                    preset.SetAttribute("subformat", cbSubFormat.Text);
                    preset.SetAttribute("progressive", BoolToVal((bool)cbProgressive.IsChecked));
                    preset.SetAttribute("optimize", BoolToVal((bool)cbOptimize.IsChecked));
                    break;
                case "PNG 24":
                    preset.SetAttribute("transparency", BoolToVal((bool)cbTransparency.IsChecked));
                    preset.SetAttribute("interlaced", BoolToVal((bool)cbInterlaced.IsChecked));
                    break;
            }

            return preset;
        }



        private void bAddPreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var w = new InputBox();
                w.ShowDialog();
                if (Ui.newName == "") return;

                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var root = xDoc.ChildNodes[1];

                int increase = Convert.ToInt32(root.Attributes["increase"].Value);
                var preset = CreateNode(xDoc, increase.ToString(CultureInfo.InvariantCulture), Ui.newName);

                root.AppendChild(preset);
                var newIncrease = (increase + 1).ToString(CultureInfo.InvariantCulture);
                root.Attributes["increase"].Value = newIncrease;

                xDoc.Save(Ui.sPath);
                LoadPresets();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }


        



        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (cbPresets.Items.Count < 1) return;
                var selItem = (ComboBoxItem)cbPresets.SelectedItem;

                if (selItem == null)
                {
                    MessageBox.Show("No selected preset.");
                    return;
                }

                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var preset = xDoc.SelectSingleNode("//Preset[@pid = \"" + selItem.Tag.ToString() + "\"]");

                if (preset != null)
                {
                    var pName = preset.Attributes["title"].Value;
                    var pid = preset.Attributes["pid"].Value;
                    var newChild = CreateNode(xDoc, pid, pName);
                    xDoc.ChildNodes[1].ReplaceChild(newChild, preset);
                    xDoc.Save(Ui.sPath);
                    LoadPresets();
                    return;
                }

                MessageBox.Show("Item not found!");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }



        private void bSetFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var root = xDoc.ChildNodes[1];
                var fAttribute = root.Attributes["folder"];

                var d = new System.Windows.Forms.FolderBrowserDialog();
                d.Description = "Choose a Folder";
                if (fAttribute != null) d.SelectedPath = root.Attributes["folder"].Value;

                if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    tbGeneralFolder.Text = d.SelectedPath;
                    if (fAttribute != null) root.Attributes["folder"].Value = d.SelectedPath;
                    else
                    {
                        var at = xDoc.CreateAttribute("folder");
                        at.Value = d.SelectedPath;
                        root.Attributes.Append(at);
                    }
                    xDoc.Save(Ui.sPath);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }


        private void bDelFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var xDoc = new XmlDocument();
                xDoc.Load(Ui.sPath);
                var root = xDoc.ChildNodes[1];
                var fAttribute = root.Attributes["folder"];

                if (fAttribute != null)
                {
                    root.Attributes.Remove(fAttribute);
                    xDoc.Save(Ui.sPath);
                    tbGeneralFolder.Text = "";
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }
        


    }
}
