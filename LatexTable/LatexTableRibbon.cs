using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace LatexTable
{
    public partial class LatexTableRibbon
    {
        
        private void LatexTableRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Load Settings
            position.Text = Properties.Settings.Default.position;
            fitWidth.Checked = Properties.Settings.Default.fitWidth;
            enableCentering.Checked = Properties.Settings.Default.enableCentering;
            skipHidden.Checked = Properties.Settings.Default.skipHidden;
            hasCaption.Checked = Properties.Settings.Default.hasCaption;
            hasLabel.Checked = Properties.Settings.Default.hasLabel;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var selectedRange = LatexTable.Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            var selectedAreas = selectedRange.Areas as Microsoft.Office.Interop.Excel.Areas;

            bool enable_hide_skip = skipHidden.Checked;

            RangeConvert rc = new RangeConvert(selectedAreas[1], enable_hide_skip);
            Tabular tab = new Tabular(rc);
            Table tb = new Table();

            tb.has_centering = enableCentering.Checked;
            tb.has_caption = hasCaption.Checked;
            tb.caption_content = Caption.Text;

            tb.has_label = hasLabel.Checked;
            tb.label_content = Label.Text;

            tb.resize = fitWidth.Checked;
            tb.position = position.Text;


            Clipboard.SetText(string.Join("\n", tb.Create_table(tab)));
            // MessageBox.Show(tb_buff);
        }

        private void SaveFileButton_Click(object sender, RibbonControlEventArgs e)
        {

            var selectedRange = LatexTable.Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            var selectedAreas = selectedRange.Areas as Microsoft.Office.Interop.Excel.Areas;

            bool enable_hide_skip = !skipHidden.Checked;

            RangeConvert rc = new RangeConvert(selectedAreas[1], enable_hide_skip);
            Tabular tab = new Tabular(rc);
            Table tb = new Table()
            {
                has_centering = enableCentering.Checked,
                has_caption = hasCaption.Checked,
                caption_content = Caption.Text,
                has_label = hasLabel.Checked,
                label_content = Label.Text,
                resize = fitWidth.Checked,
                position = position.Text
            };

            // Generate SaveFileDialog
            SaveFileDialog sa = new SaveFileDialog();
            sa.Title = "Save Table as File";
            sa.FileName = @"table.tex";
            sa.Filter = "Latex File(*.tex)|*.tex|All Files(*.*)|*.*";
            sa.FilterIndex = 1;

            // Show Dialog
            DialogResult result = sa.ShowDialog();

            if (result == DialogResult.OK)
            {
                string fileName = sa.FileName;
                var writer = new System.IO.StreamWriter(fileName, false);
                writer.WriteLine(string.Join("\n", tb.Create_table(tab)));
                writer.Close();
            }
            else if (result == DialogResult.Cancel)
            {
            }
        }

        private void position_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.position = position.Text;
        }

        private void fitWidth_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.fitWidth = fitWidth.Checked;
        }

        private void enableCentering_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.enableCentering = enableCentering.Checked;
        }

        private void skipHidden_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.skipHidden = skipHidden.Checked;
        }

        private void hasCaption_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.hasCaption = hasCaption.Checked;
        }

        private void hasLabel_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.hasLabel = hasLabel.Checked;
        }
    }
}
