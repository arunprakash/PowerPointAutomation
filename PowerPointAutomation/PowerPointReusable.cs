using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Runtime.InteropServices;

namespace PowerPointAutomation
{
    class PowerPointReusable
    {
        Application app;
        Presentation presentation;
        Slide slide;

        public void ReadPowerPoint(string source)
        {
            app = new Application();
            presentation = app.Presentations.Open(source, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
        }

        public Presentation copySlide(int rowCount)
        {
            slide = presentation.Slides[1];
            slide.Copy();
            for (int i = 1; i <= rowCount; i++)
            {
                presentation.Slides.Paste(1);
            }
            return presentation;
        }

        public void save(String filename)
        {
            presentation.SaveAs(filename);
        }

        //Sample units
        public static void editExisting()
        {
            string source = "D:\\Sample.pptx";
            string fileName = System.IO.Path.GetFileNameWithoutExtension(source);
            string filePath = System.IO.Path.GetDirectoryName(source);

            Application pa = new Application();
            Presentation pp = pa.Presentations.Open(source,
            MsoTriState.msoTrue,
            MsoTriState.msoFalse,
            MsoTriState.msoFalse);

            string pps = "";

            Slide sa = pp.Slides[1];
            sa.Copy();
            pp.Slides.Paste(1);

            foreach (Slide slide in pp.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var textFrame = shape.TextFrame;
                        if (textFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = textFrame.TextRange;
                            pps = textRange.Text.ToString();
                            if (pps.Contains("Heading"))
                            {
                                textRange.Text = "Changed Heading";
                            }
                            if (pps.Contains("Title Box"))
                            {
                                textRange.Text = "Changed Tiltle";
                            }
                            if (pps.Contains("Contents"))
                            {
                                textRange.Text = "Changed Contents1\nChanged Contents2";
                            }
                            Console.WriteLine(pps);
                        }
                    }

                }
            }
            pp.SaveAs("D:\\written.pptx");
            pa.Quit();
        }

        //Sample Units
        public static void createNew()
        {
            string pictureFileName = "D:\\test.png";

            Application pptApplication = new Application();

            Slides slides;
            _Slide slide;
            TextRange objText;

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            // Create new Slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            // Add title
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = "test";
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = "Content goes here\nYou can add text\nItem 3";

            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
            slide.Shapes.AddPicture(pictureFileName, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);

            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Test";

            pptPresentation.SaveAs(@"d:\\test.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Close();
            pptApplication.Quit();
        }

        public void closePowerPoint(Presentation presentation)
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(presentation);
            Marshal.ReleaseComObject(slide);

            //close and release
            //presentation.Close();
            //Marshal.ReleaseComObject(presentation);

            //quit and release
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
    }
}
