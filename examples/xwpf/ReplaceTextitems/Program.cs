using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.XWPF.UserModel;
using System.Reflection;

namespace DocxExample
{
    class Program
    {
        static void Insert(IBody doc, string search, string replace)
        {
            foreach (XWPFParagraph p in doc.Paragraphs)
            {
                PositionInParagraph pos = new PositionInParagraph(0, 0, 0);
                TextSegement segment = p.SearchText(search, pos);
                while (segment != null)
                {
                    for (int i = segment.BeginRun + 1; i <= segment.EndRun; i++)
                    {
                        string r = "";
                        if (i == segment.EndRun)
                        {
                            r = p.Runs[i].GetText(0);
                            r = r.Substring(segment.EndChar + 1);
                        }
                        p.Runs[i].SetText(r, 0);
                    }
                    p.Runs[segment.BeginRun].SetText(replace, segment.BeginPos.Char);
                    segment = p.SearchText(search, segment.EndPos);
                }
            }
        }

        static XWPFTable Insert(XWPFDocument doc, string search, int rows, int cols)
        {
            foreach (XWPFParagraph p in doc.Paragraphs)
            {
                PositionInParagraph pos = new PositionInParagraph(0, 0, 0);
                TextSegement segment = p.SearchText(search, pos);
                while (segment != null)
                {
                    return doc.ConvertParagraphToTable(p, rows, cols);
                }
            }
            return null;
        }

        static void Insert(XWPFDocument doc, string search, Stream data, PictureType type, string filename, int width, int height)
        {
            foreach (XWPFParagraph p in doc.Paragraphs)
            {
                PositionInParagraph pos = new PositionInParagraph(0, 0, 0);
                TextSegement segment = p.SearchText(search, pos);
                while (segment != null)
                {
                    for (int i = segment.BeginRun; i <= segment.EndRun; i++)
                    {
                        p.Runs[i].SetText("", 0);
                    }
                    p.Runs[segment.BeginRun].AddPicture(data, (int)type, filename, width, height);
                    return;
                }
            }
        }

        static void Main(string[] args)
        {

            using (Stream file = Assembly.GetExecutingAssembly().GetManifestResourceStream("DocxExample.files.test.docx"))
            {
                Type t = file.GetType();
                XWPFDocument doc = new XWPFDocument(file);

                Program.Insert(doc, "[#head1]", "Lorem Ipsum");

                Program.Insert(doc, "[#value1]", "10.000,00€");
                Program.Insert(doc, "[#value2]", "Stieleiche");


                foreach (XWPFHeader h in doc.HeaderList)
                {
                    Program.Insert(h, "[#author]", "Steffen jäckel | ARC-GREENLAB GmbH | 2014");
                    Program.Insert(h, "[#datum]", "Heute");
                }

                foreach (XWPFFooter h in doc.FooterList)
                {
                    Program.Insert(h, "[#author]", "Steffen jäckel | ARC-GREENLAB GmbH | 2014");
                    Program.Insert(h, "[#date]", new DateTime().ToShortDateString() );
                }


                XWPFTable table = Program.Insert(doc, "[#table1]", 3, 6);
                table.Width = 5000;///
                int temp = NPOI.Util.Units.ToEMU(450);
                table.Rows[0].GetCell(0).SetText("First");
                table.Rows[0].GetCell(1).SetText("Second");
                table.Rows[0].GetCell(2).SetText("Third");
                table.Rows[0].GetCell(3).SetText("Fourth");
                table.Rows[0].GetCell(4).SetText("Fifth");
                table.Rows[0].GetCell(5).SetText("Sixth");
                for (int r = 1; r < 3; r++)
                {
                    for (int c = 0; c < 6; c++)
                    {
                        table.Rows[r].GetCell(c).SetText("0");
                    }
                }

                using (Stream img = Assembly.GetExecutingAssembly().GetManifestResourceStream("DocxExample.files.img.png"))
                {
                    Program.Insert(doc, "[#img1]", img, PictureType.PNG, "img.png", NPOI.Util.Units.ToEMU(450), NPOI.Util.Units.ToEMU(300));
                }

                FileStream sw = File.Create("export.docx");
                doc.Write(sw);
                sw.Close();
            }
        }
    }
}
