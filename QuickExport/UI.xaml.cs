using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;
using Corel.Interop.VGCore;
using Page = Corel.Interop.VGCore.Page;

namespace QuickExport
{
    public partial class Ui : UserControl
    {

        public static Corel.Interop.VGCore.Application dApp = null;
        public const string mName = "QuickExport";
        public static string mVer = "1.2";
        public static string mDate = "14.04.2014";
        public static string mWebSite = "http://cdrpro.ru";
        public static string mEmail = "sancho@cdrpro.ru";

        private string _shTitle = "QuickExport #macro for #CorelDRAW";
        private string _shUrl = "http://cdrpro.ru/";

        public static string sPath;
        public static string newName = "";
        


        public Ui() { InitializeComponent(); }
        public Ui(object app)
        {
            try
            {
                InitializeComponent();
                dApp = (Corel.Interop.VGCore.Application)app;

                var uFolderPath = Environment.GetEnvironmentVariable("APPDATA") + @"\Corel\" + mName;
                if (!Directory.Exists(uFolderPath)) Directory.CreateDirectory(uFolderPath);
                sPath = uFolderPath + @"\Presets.xml";

                if (!File.Exists(sPath))
                {
                    var xDoc = new XmlDocument();
                    var dec = xDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
                    xDoc.AppendChild(dec);
                    var root = xDoc.CreateElement("Presets");
                    root.SetAttribute("increase", "1");
                    xDoc.AppendChild(root);
                    xDoc.Save(sPath);
                }
                LoadPresets();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), mName, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        public void LoadPresets()
        {
            try
            {
                btnContext.Items.Clear();

                var xDoc = new XmlDocument();
                xDoc.Load(sPath);
                if (xDoc.ChildNodes.Count == 0) return;

                MenuItem newMenuItem = null;

                if (xDoc.ChildNodes[1].ChildNodes.Count == 0)
                {
                    newMenuItem = new MenuItem { Header = "Presets not found", IsEnabled = false };
                    btnContext.Items.Add(newMenuItem);
                }
                else
                {
                    foreach (XmlNode n in xDoc.ChildNodes[1].ChildNodes)
                    {
                        newMenuItem = new MenuItem { Header = n.Attributes["title"].Value, Tag = n.Attributes["pid"].Value };
                        newMenuItem.Click += new RoutedEventHandler(BtnContextDo);
                        btnContext.Items.Add(newMenuItem);
                    }
                }

                btnContext.Items.Add(new Separator());

                newMenuItem = new MenuItem { Header = "About " + mName + "...", Tag = "about" };
                newMenuItem.Click += new RoutedEventHandler(ShowAbout);
                btnContext.Items.Add(newMenuItem);
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }




        private void BtnBorderMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var w = new wSettings();
                var wih = new System.Windows.Interop.WindowInteropHelper(w);
                wih.Owner = (IntPtr)dApp.AppWindow.Handle;
                w.ShowDialog();
                LoadPresets();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }



        private void ShowAbout(object sender, RoutedEventArgs e)
        {
            try
            {
                var w = new wAbout();
                var wih = new System.Windows.Interop.WindowInteropHelper(w);
                wih.Owner = (IntPtr)dApp.AppWindow.Handle;
                w.ShowDialog();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }



        private bool ValToBool(string val) { return val == "1" ? true : false; }



        private void BtnContextDo(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dApp.Documents.Count == 0) return;

                var mi = (MenuItem) sender;
                var pid = mi.Tag.ToString();

                var xDoc = new XmlDocument();
                xDoc.Load(sPath);
                var root = xDoc.ChildNodes[1];

                var preset = xDoc.SelectSingleNode("//Preset[@pid = \"" + pid + "\"]");

                if (preset != null)
                {
                    var d = dApp.ActiveDocument;

                    string filePath = "";
                    if (root.Attributes["folder"] != null) filePath = root.Attributes["folder"].Value + @"\";
                    else
                    {
                        filePath = d.FilePath;
                        if (filePath == "") filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\";
                    }
                    filePath += d.Name;

                    cdrFilter filter;
                    string ext;
                    var opt = new StructExportOptions();

                    var aat = cdrAntiAliasingType.cdrNoAntiAliasing;
                    if (preset.Attributes["antialiased"].Value == "1") aat = cdrAntiAliasingType.cdrNormalAntiAliasing;

                    var imgType = cdrImageType.cdrRGBColorImage;
                    switch (preset.Attributes["colormode"].Value)
                    {
                        case "CMYK Color":
                            imgType = cdrImageType.cdrCMYKColorImage;
                            break;
                        case "Grayscale":
                            imgType = cdrImageType.cdrGrayscaleImage;
                            break;
                    }

                    opt.Overwrite = true;
                    opt.AntiAliasingType = aat;
                    opt.ImageType = imgType;
                    opt.MaintainAspect = true;
                    opt.UseColorProfile = true;
                    opt.ResolutionX = Convert.ToInt32(preset.Attributes["resolution"].Value);
                    opt.ResolutionY = Convert.ToInt32(preset.Attributes["resolution"].Value);
                    opt.Transparent = false;

                    var format = preset.Attributes["format"].Value;
                    if (format == "Jpeg")
                    {
                        filter = cdrFilter.cdrJPEG;
                        ext = ".jpg";
                        opt.Compression = cdrCompressionType.cdrCompressionJPEG;
                    }
                    else
                    {
                        filter = cdrFilter.cdrPNG;
                        ext = ".png";
                        if (preset.Attributes["transparency"].Value == "1") opt.Transparent = true;
                    }

                    switch (preset.Attributes["range"].Value)
                    {
                        case "Selection":
                            if (dApp.ActiveSelectionRange.Count == 0)
                            {
                                MessageBox.Show("No selection");
                                return;
                            }
                            ExportImage(d, preset, filePath + ext, filter, cdrExportRange.cdrSelection, opt);
                            break;

                        case "Active page":
                            if (dApp.ActivePage.Shapes.Count == 0)
                            {
                                MessageBox.Show("No shapes");
                                return;
                            }
                            ExportImage(d, preset, filePath + @"_" + d.ActivePage.Index.ToString(CultureInfo.InvariantCulture) + ext, filter, cdrExportRange.cdrCurrentPage, opt);
                            break;

                        case "All pages":
                            foreach (Page p in d.Pages)
                            {
                                p.Activate();
                                ExportImage(d, preset, filePath + @"_" + p.Index.ToString(CultureInfo.InvariantCulture) + ext, filter, cdrExportRange.cdrCurrentPage, opt);
                            }
                            break;
                    }

                    return;
                }

                MessageBox.Show("Preset not found");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }



        private bool ExportImage(Document d, XmlNode preset, string filePath, cdrFilter filter, cdrExportRange range, StructExportOptions opt)
        {
            try
            {

                Corel.Interop.VGCore.ICorelExportFilter exp;
                exp = d.ExportEx(filePath, filter, range, opt);
                object[] param = new Object[1];

                if (preset.Attributes["format"].Value == "Jpeg")
                {
                    param[0] = ValToBool(preset.Attributes["progressive"].Value);
                    exp.GetType().InvokeMember("Progressive", BindingFlags.SetProperty, null, exp, param);

                    param[0] = ValToBool(preset.Attributes["optimize"].Value);
                    exp.GetType().InvokeMember("Optimized", BindingFlags.SetProperty, null, exp, param);

                    if (preset.Attributes["subformat"].Value == "Standard (4:2:2)") param[0] = 0;
                    else param[0] = 1;
                    exp.GetType().InvokeMember("SubFormat", BindingFlags.SetProperty, null, exp, param);

                    param[0] = 100 - Convert.ToInt32(preset.Attributes["quality"].Value);
                    exp.GetType().InvokeMember("Compression", BindingFlags.SetProperty, null, exp, param);

                    param[0] = Convert.ToInt32(preset.Attributes["blur"].Value);
                    exp.GetType().InvokeMember("Smoothing", BindingFlags.SetProperty, null, exp, param);
                }
                else
                {
                    param[0] = ValToBool(preset.Attributes["interlaced"].Value);
                    exp.GetType().InvokeMember("Interlaced", BindingFlags.SetProperty, null, exp, param);
                }

                exp.Finish();
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        }


    }
}
