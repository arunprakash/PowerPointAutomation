using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;

namespace PowerPointAutomation
{
    public partial class PowerPointAutomation : Form
    {


        public PowerPointAutomation()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReusable excel = new ExcelReusable();
            Range xlRange = excel.ReadExcel("d:\\Sample.xlsx");

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            PowerPointReusable powerpoint = new PowerPointReusable();
            powerpoint.ReadPowerPoint("D:\\Sample.pptx");

            Presentation presentation = powerpoint.copySlide(rowCount-1);
            
            //foreach (Slide slide in presentation.Slides)

           for (int i = 1; i <= rowCount; i++)
           {
                Slide slide = presentation.Slides[i];

                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var textFrame = shape.TextFrame;
                        if (textFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = textFrame.TextRange;
                            string text = textRange.Text.ToString();
                            if (text.Contains("Heading"))
                            {
                                textRange.Text = xlRange.Cells[i, 1].Value2.ToString();
                            }
                            if (text.Contains("Title Box"))
                            {
                                textRange.Text = xlRange.Cells[i, 2].Value2.ToString();
                            }
                            if (text.Contains("Contents"))
                            {
                                textRange.Text = xlRange.Cells[i, 3].Value2.ToString();
                            }
                        }
                    }

                }
           }
            powerpoint.save("D:\\written.pptx");
            powerpoint.closePowerPoint(presentation);
            excel.closeExcel();
            Console.WriteLine("Done");
        }
    }
}
