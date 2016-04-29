using Microsoft.Win32;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Xml.Linq;
using System;
using System.Windows.Media;
using System.Windows.Controls;

namespace TranslationImporter
{
    /*
        TODO:   Look into copy pasted text in the settings fields.
        TODO:   Add export functionality - resource to excel.
        TODO:   Add import/export result statistics.
        TODO:   Use watermarks in the textboxes.
        TODO:   Add ability to choose whether to delete all existing contents before importing/exporting.
    */

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    { 
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Event Handling

        protected void btnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files (*.xls, *.xlsx, *.xlsm, *.xlsb, *.xltx, *.xltm, *.xlt)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xlt";
            bool? userClickedOK = ofd.ShowDialog();

            if (userClickedOK.HasValue && userClickedOK.Value)
            {
                tbExcelFile.Text = ofd.FileName;
            }
        }

        protected void btnBrowseResx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Resource Files (*.resx)|*.resx";
            bool? userClickedOK = ofd.ShowDialog();

            if (userClickedOK.HasValue && userClickedOK.Value)
            {
                tbResxFile.Text = ofd.FileName;
            }
        }

        protected void lbShowStats_Click(object sender, EventArgs e)
        {
            // TODO:
            // Create a panel that shows the stats of the import.
            // 1. How many keys added
            // 2. How many keys updated
            // 3. How many keys removed
            // 4. Blank keys
        }

        protected void tbExcelFile_GotKeyboardFocus(object sender, EventArgs e)
        {
            // TODO: Add watermarks to the textboxes
            if (tbExcelFile.Text == AppResources.ExcelFileTextboxPlaceholder)
            {
                tbExcelFile.Text = "";
            }

            tbExcelFile.FontStyle = FontStyles.Normal;
        }

        protected void tbExcelFile_LostKeyboardFocus(object sender, EventArgs e)
        {
            if (tbExcelFile.Text.Length < 1)
            {
                tbExcelFile.Text = AppResources.ExcelFileTextboxPlaceholder;
                tbExcelFile.FontStyle = FontStyles.Italic;
            }
            else
            {
                tbExcelFile.FontStyle = FontStyles.Normal;
            }
        }

        protected void tbResxFile_GotKeyboardFocus(object sender, EventArgs e)
        {
            if (tbResxFile.Text == AppResources.ResourceFileTextboxPlaceholder)
            {
                tbResxFile.Text = "";
            }

            tbResxFile.FontStyle = FontStyles.Normal;
        }

        protected void tbResxFile_LostKeyboardFocus(object sender, EventArgs e)
        {
            if (tbResxFile.Text.Length < 1)
            {
                tbResxFile.Text = AppResources.ResourceFileTextboxPlaceholder;
                tbResxFile.FontStyle = FontStyles.Italic;
            }
            else
            {
                tbResxFile.FontStyle = FontStyles.Normal;
            }
        }

        protected void btnImport_Click(object sender, RoutedEventArgs e)
        {
            ClearMessage();

            FileInfo excelFile = IsValidExcelFile(tbExcelFile.Text);
            FileInfo resxFile = IsValidResourceFile(tbResxFile.Text);

            if (excelFile != null && resxFile != null)
            {
                using (ExcelPackage ep = new ExcelPackage(excelFile))
                {
                    Dictionary<string, string> labelsAndText = new Dictionary<string, string>();
                    ExcelWorksheet ws = null;
                    int worksheetIndex = 0;

                    if (int.TryParse(tbWorksheetIndex.Text, out worksheetIndex))
                    {
                        ws = ep.Workbook.Worksheets[int.Parse(tbWorksheetIndex.Text)];

                        // Get all of the data labels.
                        var cells = ws.Cells[tbDataLabelColumn.Text + ":" + tbDataLabelColumn.Text].AsEnumerable<ExcelRangeBase>();

                        // Skip the first row if the user chose to
                        // do so. It may just be column headers.
                        if (cbSkipFirstRow.Text == "Yes")
                            cells = cells.Skip(1);

                        // Get the corresponding text for each of the
                        // labels and build a list of the results.
                        foreach (ExcelRangeBase dataLabelCell in cells)
                        {
                            int row = dataLabelCell.Start.Row;
                            ExcelRangeBase translatedTextCell = ws.Cells[tbTranslatedTextColumn.Text + "" + row];

                            labelsAndText.Add(dataLabelCell.Text, translatedTextCell.Text);
                        }

                        // Only proceed if there are any labels in the worksheet.
                        if (labelsAndText.Count > 0)
                        {
                            XDocument resourceFile = XDocument.Load(resxFile.FullName);
                            int updatedCount = 0;
                            int insertedCount = 0;

                            foreach (KeyValuePair<string, string> kvp in labelsAndText)
                            {
                                var existingResource = resourceFile.Descendants("data")
                                                                   .Where(x => x.Attribute("name").Value == kvp.Key);

                                // If there are any existing labels in the resource
                                // file with this name, update the value. Otherwise,
                                // create a new element and add it to the file.
                                if (existingResource.Any())
                                {
                                    resourceFile.Descendants("data")
                                                .Where(x => x.Attribute("name").Value == kvp.Key)
                                                .Single()
                                                .SetElementValue("value", kvp.Value);

                                    updatedCount++;
                                }
                                else
                                {
                                    XElement element = new XElement("data",
                                        new XAttribute("name", kvp.Key),
                                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                                        new XElement("value", kvp.Value)
                                    );

                                    resourceFile.Root.Add(element);

                                    insertedCount++;
                                }
                            }

                            // Save all of the changes.
                            resourceFile.Save(resxFile.FullName);

                            SuccessMessage("Import successful");
                        }
                        else
                        {
                            ErrorMessage("No labels were found. Are the settings below correct?");
                        }
                    }
                    else
                    {
                        ErrorMessage("The worksheet index you entered isn't an integer.");
                    }
                }
            }
        }

        private void tbWorksheetIndex_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Only numeric
            int i;
            e.Handled = !int.TryParse(e.Text, out i);
        }

        private void tbDataLabelColumn_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Only alpha
            e.Handled = !e.Text.All(Char.IsLetter);
        }

        private void tbTranslatedTextColumn_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Only alpha
            e.Handled = !e.Text.All(Char.IsLetter);
        }

        #endregion

        #region Utility Methods

        private FileInfo IsValidExcelFile(string fileName)
        {
            // TODO: Remove the file name text check by adding watermark to the textbox
            FileInfo file = null;
            if (fileName.Length < 1 || fileName == AppResources.ExcelFileTextboxPlaceholder)
            {
                lblExcelMessage.Text = "Please choose an excel file.";
                lblExcelMessage.Height = double.NaN;
            }
            else if (!IsExcelFile(fileName))
            {
                lblExcelMessage.Text = "Please use one of these formats: *.xls, *.xlsx, *.xlsm, *.xlsb, *.xltx, *.xltm, *.xlt";
                lblExcelMessage.Height = double.NaN;
            }
            else
            {
                file = new FileInfo(fileName);
                if (!file.Exists)
                {
                    lblExcelMessage.Text = "Please choose an excel file that exists.";
                    lblExcelMessage.Height = double.NaN;
                    file = null;
                }
                else
                {
                    lblExcelMessage.Text = "";
                    lblExcelMessage.Height = 0;
                }
            }

            return file;
        }

        private FileInfo IsValidResourceFile(string fileName)
        {
            // TODO: Remove the file name text check by adding watermark to the textbox
            FileInfo file = null;
            if (fileName.Length < 1 || fileName == AppResources.ResourceFileTextboxPlaceholder)
            {
                lblResourceMessage.Text = "Please choose a resource file.";
                lblResourceMessage.Height = double.NaN;
            }
            else if (!IsResourceFile(fileName))
            {
                lblResourceMessage.Text = "Please use the *.resx format.";
                lblResourceMessage.Height = double.NaN;
            }
            else
            {
                file = new FileInfo(fileName);
                if (!file.Exists)
                {
                    lblResourceMessage.Text = "Please choose a resource file that exists.";
                    lblResourceMessage.Height = double.NaN;
                    file = null;
                }
                else
                {
                    lblResourceMessage.Text = "";
                    lblResourceMessage.Height = 0;
                }
            }

            return file;
        }

        private bool IsExcelFile(string fileName)
        {
            if (fileName.EndsWith(".xlsx")
                || fileName.EndsWith(".xlsm")
                || fileName.EndsWith(".xlsb")
                || fileName.EndsWith(".xltx")
                || fileName.EndsWith(".xltm")
                || fileName.EndsWith(".xls")
                || fileName.EndsWith(".xlt"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool IsResourceFile(string fileName)
        {
            if (fileName.EndsWith(".resx"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void ErrorMessage(string text)
        {
            lblMessage.Text = text;
            lblMessage.Foreground = Brushes.Red;
            Grid.SetColumnSpan(lblMessage, 2);
        }

        private void SuccessMessage(string text)
        {
            lblMessage.Text = text;
            lblMessage.Foreground = Brushes.Green;

            // TODO: Show stats button
            //lbShowStats.Visibility = System.Windows.Visibility.Visible;
        }

        private void ClearMessage()
        {
            lblMessage.Text = "";
            //lbShowStats.Visibility = System.Windows.Visibility.Hidden;
        }

        #endregion
    }
}
