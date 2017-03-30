using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RTFLibDemo
{
    using GDF;
    using Properties;
    using RTF;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RTFBuilderbase sb = new RTFBuilder();

            BuilderCode(sb);

       

            this.richTextBox1.Rtf = sb.ToString();
        }

        private void BuilderCode(RTFBuilderbase sb)
        {
            sb.AppendLine("AppendLine Basic Text");
            sb.Append("append text1").Append("append text2").Append("append text3").Append("append text4").AppendLine();
            sb.FontStyle(FontStyle.Bold).AppendLine("Bold");
            sb.FontStyle(FontStyle.Italic).AppendLine("Italic");
            sb.FontStyle(FontStyle.Strikeout).AppendLine("Strikeout");
            sb.FontStyle(FontStyle.Underline).AppendLine("Underline");
            sb.FontStyle(FontStyle.Bold | FontStyle.Italic | FontStyle.Strikeout | FontStyle.Underline).AppendLine("Underline/Bold/Italic/Underline");
            sb.ForeColor(KnownColor.Red).AppendLine("ForeColor Red");
            sb.BackColor(KnownColor.Yellow).AppendLine("BackColor Yellow");

            sb.ForeColor(KnownColor.Red).BackColor(KnownColor.Yellow).AppendLine("ForeColor Red , BackColor Yellow");

            sb.AppendLine("1. append 2 lines - First Line \r\n2. Second Line (correctly appending rtf codes for secondline)");
            sb.Font(RTFFont.Georgia).AppendLine("Change to Georgia Font!");
            sb.Font(RTFFont.Consolas).AppendLine("Change to Consolas Font!");
            sb.Font(RTFFont.Garamond).AppendLine("Change to Garamond Font!");
            sb.Font(RTFFont.MSSansSerif).AppendLine("Change to MSSansSerif Font!)");
            sb.Font(RTFFont.Arial).AppendLine("Change to Arial Font!(default)");
            sb.Font(RTFFont.ArialBlack).AppendLine("Change to ArialBlack Font!");

            sb.FontSize(30).Font(RTFFont.ArialBlack).AppendLine("Change to ArialBlack Font Size 30");

            //Commit Format changes
            sb.FontSize(20).Font(RTFFont.ArialBlack);

            using (sb.FormatLock())
            {
                sb.AppendLine("FormatLock to ArialBlack Font Size 20");
                sb.AppendLine("FormatLock to ArialBlack Font Size 20");

            }
            sb.AppendLine("Inserting Image");
            sb.InsertImage(Resources.Complications);
            sb.AppendLine("Inserted Image");

               sb.AppendLine("Creating Table");
               sb.AppendPage();
                        sb.AppendPara();
                        sb.Reset();
    
               AddRow1(sb, "Row 1 Cell 1", "Row 1 Cell 2", "Row 1 Cell 3");
               AddRow1(sb, "Row 2 Cell 1", "Row 2 Cell 2", "Row 2 Cell 3");
               AddRow1(sb, "Row 3 Cell 1", "Row 3 Cell 2", "Row 3 Cell 3");
               AddRow1(sb, "Row 4 Cell 1", "Row 4 Cell 2", "Row 4 Cell 3");
               AddRow1(sb, "Row 5 Cell 1", "Row 5 Cell 2", "Row 5 Cell 3");
      
               sb.AppendLine("Inserting cell images");
               sb.AppendPara();
                 sb.Reset();
                 sb.AppendPara();
                 sb.Reset();
               AddRow2(sb, "Row 5 Cell 1", "Row 5 Cell 2", "Row 5 Cell 3");
               sb.AppendPara();
               sb.AppendLine("Inserting MultiLine Cells --> Fails to display correctly in RichTextBox");
               sb.Reset();
               sb.AppendPara();
               AddRow1(sb, "Row 5\r\n Cell 1", "Row 5 \r\n Cell 2", "Row 5 \r\n Cell 3");
               sb.Reset();
               sb.AppendPara();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Size s = gdfDisplayBox1.ClientSize;
            s.Height -= 31;


            GDFPageManager manager = new GDFPageManager(s, SystemColors.Window);

            GDFBuilder sb = new GDFBuilder(manager);
            BuilderCode(sb);
            gdfDisplayBox1.SetPages(manager.Pages);


        }

        /// <summary>
        /// Adds the row1.
        /// </summary>
        /// <param name="sb">The sb.</param>
        /// <param name="cellContents">The cell contents.</param>
        private void AddRow1(RTFBuilderbase sb, params string[] cellContents)
        {
            Padding p = new Padding { All = 50 };
            RTFRowDefinition rd = new RTFRowDefinition(88, RTFAlignment.TopLeft, RTFBorderSide.Default, 15, SystemColors.WindowText, p);
            RTFCellDefinition[] cds = new RTFCellDefinition[cellContents.Length];
            for (int i = 0; i < cellContents.Length; i++)
            {
                cds[i] = new RTFCellDefinition(88 / cellContents.Length, RTFAlignment.TopLeft, RTFBorderSide.Default, 15, Color.Blue, Padding.Empty);
            }
            int pos = 0;
            foreach (RTFBuilderbase item in sb.EnumerateCells(rd, cds))
            {
                item.ForeColor(KnownColor.Blue).FontStyle(FontStyle.Bold | FontStyle.Underline);
                item.BackColor(Color.Yellow);
                item.Append(cellContents[pos++]);
            }
        }
        /// <summary>
        /// utility function to facilitate insertion of cellContents within a row
        /// </summary>
        /// <param name="sb">The sb.</param>
        /// <param name="cellContents">The cell contents.</param>
        private void AddRow2(RTFBuilderbase sb, params string[] cellContents)
        {
            Padding p = new Padding { All = 50 };//ignored
            // create RTFRowDefinition
            RTFRowDefinition rd = new RTFRowDefinition(88, RTFAlignment.TopLeft, RTFBorderSide.Default, 15, SystemColors.WindowText, p);
            // Create RTFCellDefinitions
            RTFCellDefinition[] cds = new RTFCellDefinition[cellContents.Length];
            for (int i = 0; i < cellContents.Length; i++)
            {
                cds[i] = new RTFCellDefinition(88 / cellContents.Length, RTFAlignment.TopLeft, RTFBorderSide.Default, 15, Color.Blue, Padding.Empty);
            }
            int pos = 0;
            // enumerate over cells
            // each cell 
            foreach (RTFBuilderbase item in sb.EnumerateCells(rd, cds))
            {
                item.InsertImage(Resources.Complications);
                item.ForeColor(KnownColor.Blue).FontStyle(FontStyle.Bold | FontStyle.Underline);
                item.BackColor(Color.Yellow);
                item.Append(cellContents[pos++]);
            }
        }

        /// <summary>
        /// Merges the output of one RtfBuilder into another
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void button3_Click(object sender, EventArgs e)
        {
            RTFBuilderbase sb = new RTFBuilder();

            BuilderCode(sb);
            RTFBuilderbase sb2 = new RTFBuilder();
            BuilderCode(sb2);
            sb.AppendRTFDocument(sb2.ToString());
            this.richTextBox2.Rtf = sb.ToString();
        }

        /// <summary>
        /// Merges the contents of 2 RichTextBoxes
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void button4_Click(object sender, EventArgs e)
        {
            RTFBuilderbase sb = new RTFBuilder();
            BuilderCode(sb);
            this.richTextBox3.Rtf = sb.ToString();
            this.richTextBox4.Rtf = sb.ToString();
            sb.Clear();
            sb.AppendRTFDocument(this.richTextBox3.Rtf);
            sb.AppendRTFDocument(this.richTextBox4.Rtf);
            this.richTextBox2.Rtf = sb.ToString();
        }
    }
}
