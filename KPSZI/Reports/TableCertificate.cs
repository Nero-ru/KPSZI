using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace KPSZI.Reports
{
    class TableCertificate : StageReport
    {
        protected override Encoding htmlEncoding { get => Encoding.UTF8; }

        public TableCertificate(MainForm mainForm, string template)
            : base(mainForm, template)
        {

        }

        public override void Parce(string pathHTML)
        {
            throw new NotImplementedException();
        }

        public override void ReportToWord(string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null)
        {
            app = new Word.Application();
            doc = app.Documents.Add(Environment.CurrentDirectory + "/" + nameWord);
            paragraph = doc.Paragraphs.Add();
            app.Visible = true;

            int countSZI = mf.lvReportCertificates.Items.Count;
            Word.Table table = CreateStandartTable(paragraph.Range, countSZI, 2, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
        }
    }
}
