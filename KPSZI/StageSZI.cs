using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace KPSZI
{
    class StageSZI:Stage
    {
        protected override ImageList imageListForTabPage { get; set; }
        List<SZI> listOfSZIsFromDB;
        List<SZISort> listOfNeededSZISorts;
        List<RadioButton> radiobuttonsSZISorts;
        List<CheckBox> checkboxesSZISorts;

        public StageSZI(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {
            mf.btnGetRequirm.Click += BtnGetRequirm_Click;
            mf.btnGetSZIs.Click += BtnGetSZIs_Click;
            mf.btnGetMeasSZIs.Click += BtnGetMeasSZIs_Click;
            mf.btnExportTP.Click += BtnExportTP_Click;
            mf.tabControlSZIs.Selecting += TabControlSZIs_Selecting;
        }

        

        protected override void initTabPage()
        {
            mf.tabControlSZIs.TabPages.Clear();

            mf.dgvSZIs.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            mf.dgvSZIs.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            mf.dgvSZIs.Columns.Add("Count", @"№ п\п");
            mf.dgvSZIs.Columns[0].Width = 30;
            mf.dgvSZIs.Columns.Add("Name", "Наименование СЗИ");
            mf.dgvSZIs.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            mf.dgvSZIs.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            mf.dgvSZIs.Columns.Add("Sort", "Относится к видам");
            mf.dgvSZIs.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            mf.dgvSZIs.Columns.Add("Cert", "Сертификат");
            mf.dgvSZIs.Columns[3].Width = 120;
            //mf.dgvSZIs.Columns.Add("Class", "Класс защищенности");
            //mf.dgvSZIs.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            mf.dgvSZIs.Columns.Add("NDV", "УКО НДВ");
            mf.dgvSZIs.Columns[4].Width = 35;
            mf.dgvSZIs.Columns.Add("SVT", "СВТ");
            mf.dgvSZIs.Columns[5].Width = 35;
            mf.dgvSZIs.Columns.Add("TU", "ТУ");
            mf.dgvSZIs.Columns[6].Width = 30;


            mf.dgvMeasSZIs.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            mf.dgvMeasSZIs.Columns.Add("Count", @"№ п\п");
            mf.dgvMeasSZIs.Columns.Add("Measure", "Наименование меры");
            mf.dgvMeasSZIs.Columns.Add("SZIs", "Техническое средство");
            mf.dgvMeasSZIs.Columns[0].Width = 30;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvMeasSZIs.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            mf.dgvMeasSZIs.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        public override void saveChanges()
        {

        }

        public override void enterTabPage()
        {

        }

        #region События
        private void BtnGetRequirm_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            if (IS.listOfAllNSDMeasures.Count == 0)
            {
                if (MessageBox.Show("Не определен перечень мер защиты информации, необходимый для дальнейшней работы. \nПерейти во вкладку \"Перечень мер\"?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnMeasures");
                    return;
                }
                else
                    return;
            }

            mf.tabControlSZIs.TabPages.Clear();
            mf.tpSZItbInfo.Text = "Ниже представлены сформированные требования к средствам защиты информации согласно классу защищенности информационной системы и условиям функционирования. В соответствии с требованиям сформирован перечень видов СЗИ и их представителей. После осуществления выбора, нажмите \"СЗИ\", чтобы просмотреть информацию о них.";
            switch (IS.GISClass)
            {
                case 1: { mf.tbtpSZISVT.Text = "не ниже 5-го класса"; mf.tbtpSZISZI.Text = "не ниже 4-го класса"; mf.tbtpSZINDV.Text = "4"; break; }
                case 2: { mf.tbtpSZISVT.Text = "не ниже 5-го класса"; mf.tbtpSZISZI.Text = "не ниже 5-го класса"; mf.tbtpSZINDV.Text = "4"; break; }
                case 3: { mf.tbtpSZISVT.Text = "не ниже 5-го класса"; mf.tbtpSZISZI.Text = "не ниже 6-го класса"; mf.tbtpSZINDV.Text = "не требуется"; break; }
                default:
                    {
                        if (MessageBox.Show("Не определен класс защищенности для дальнейшней работы. \nПерейти во вкладку \"Классификация\"?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            mf.treeView.SelectedNode = mf.returnTreeNode("tnClassification");
                            return;
                        }
                        else
                            return;
                    }
            }


            mf.tabControlSZIs.TabPages.Add(mf.tpSZItpReq);
            using (KPSZIContext db = new KPSZIContext())
            {
                listOfSZIsFromDB = db.SZIs.ToList();
                listOfNeededSZISorts = new List<SZISort>();
                var listOfAllNSDMeasuresTemporary = db.GisMeasures.ToList().Intersect(IS.listOfAllNSDMeasures).ToList();
                foreach (GISMeasure gm in listOfAllNSDMeasuresTemporary)
                {
                    foreach (SZISort ss in gm.SZISorts)
                    {
                        listOfNeededSZISorts.Add(ss);
                    }
                }

                listOfNeededSZISorts = listOfNeededSZISorts.Distinct().ToList();

                GroupBox gb;
                RadioButton rb;
                CheckBox cb;
                radiobuttonsSZISorts = new List<RadioButton>();
                checkboxesSZISorts = new List<CheckBox>();

                mf.ptpSZIforSZI.Controls.Clear();

                int i = 0;
                int j = 0;
                int gbY = 5;

                foreach (SZISort szisort in listOfNeededSZISorts)
                {
                    gb = new GroupBox();
                    gb.Location = new Point(7, gbY);
                    gb.Text = szisort.Name;

                    foreach (SZI szi in szisort.SZIs)
                    {
                        if (szisort.ShortName != "САВЗ")
                        {
                            rb = new RadioButton();
                            rb.Text = szi.Name;
                            rb.Margin = new Padding(10, 5, 5, 5);
                            rb.Location = new Point(10, 17 + (17 * j));
                            rb.Size = new Size(300, 17);
                            rb.CheckedChanged += Rb_CheckedChanged;
                            radiobuttonsSZISorts.Add(rb);
                            gb.Controls.Add(rb);
                            j++;
                        }
                        else
                        {
                            cb = new CheckBox();
                            cb.Text = szi.Name;
                            cb.Margin = new Padding(10, 5, 5, 5);
                            cb.Location = new Point(10, 17 + (17 * j));
                            cb.Size = new Size(300, 17);
                            cb.CheckedChanged += Rb_CheckedChanged;
                            checkboxesSZISorts.Add(cb);
                            gb.Controls.Add(cb);
                            j++;
                        }

                        //ставим птичку на первом добавленном СЗИ
                        if (gb.Controls.Count == 1)
                        {
                            var ctrl = gb.Controls[0];
                            if (ctrl is CheckBox)
                            {
                                var _cb = (CheckBox)ctrl;
                                _cb.Checked = true;
                            }

                            if (ctrl is RadioButton)
                            {
                                var _rb = (RadioButton)ctrl;
                                _rb.Checked = true;
                            }
                        }
                    }

                    gb.Size = new Size(300, 25 + j * 17);
                    gbY += 30 + j * 17;
                    j = 0;
                    //stageTab.Controls.Add(gb);

                    string description;
                    switch (szisort.ShortName)
                    {
                        case "СЗИ от НСД": { description = "Средство защиты информации от несанкционированного доступа. Может реализовывать функции средства контроля съемных машинных носителей"; break; }
                        case "СОВ": { description = "Средство обнаружения вторжений. Может реализовывать функции межсетевых экранов, поэтому не имеет смысла выбирать отличное от него, т.к. приведет к дополнительным затратам. Требования к СОВ утверждены Приказом ФСТЭК России от 6 декабря 2011 г. № 638."; break; }
                        case "МСЭ": { description = "Межсетевой экран. Может реализовывать функции средств обнаружения вторжений, поэтому не имеет смысла выбирать отличное от него, т.к. приведет к дополнительным затратам. Требования к Межсетевым экранам утверждены Приказом ФСТЭК России от 9 февраля 2016 г. № 9."; break; }
                        case "СДЗ": { description = "Средство доверенной загрузки. Обеспечивает доверенную загрузку СВТ. Требования к СДЗ утверждены Приказом ФСТЭК России от 27 сентября 2013 г. № 119."; break; }
                        case "САВЗ": { description = "Средство антивирусной защиты. Обеспечивает защиту от компьютерных вирусов. Требования к САВЗ утверждены Приказом ФСТЭК России от 20 марта 2012 г. № 28."; break; }
                        case "СКН": { description = "Средство контроля съемных машинных носителей информации. Обеспечивает защиту от компьютерных вирусов. Требования к СКН утверждены Приказом ФСТЭК России от 28 июля 2014 г. № 87."; break; }
                        case "САНЗ": { description = "Средство анализа (контроля) защищенности информационной системы. Предназначено для анализа защищенности информации в корпоративных сетях. Специальных требований, кроме наличия сертификата ФСТЭК России, к данному виду СЗИ не устанавливается."; break; }
                        case "СРСБ": { description = "Средство регистрации событий безопасности. Предназначено для управления событиями и информацией ИБ с целью выявления инцидентов в режиме реального времени. Специальных требований, кроме наличия сертификата ФСТЭК России, к данному виду СЗИ не устанавливается."; break; }
                        case "СЗСВ": { description = "Средство защиты среды виртуализации. Предназначено для защиты среды виртуальной инфраструктуры предприятия. Специальных требований, кроме наличия сертификата ФСТЭК России, к данному виду СЗИ не устанавливается."; break; }
                        case "СРК": { description = "Средство резервного копирования и восстановления информации. Предназначено для обеспечения целостности и доступности информации. Специальных требований, кроме наличия сертификата ФСТЭК России, к данному виду СЗИ не устанавливается."; break; }

                        default: { description = "Описание отсутствует"; break; }
                    }

                    TextBox tb = new TextBox();
                    tb.Text = description;
                    tb.Multiline = true;
                    tb.TextAlign = HorizontalAlignment.Left;
                    tb.Location = new Point(gb.Location.X + gb.Size.Width + 5, gb.Location.Y + 8);
                    tb.Size = new Size(stageTab.Size.Width - 7 - gb.Size.Width - 5 - 50, gb.Size.Height - 8);
                    tb.BackColor = SystemColors.Control;
                    tb.ReadOnly = true;


                    mf.ptpSZIforSZI.Controls.Add(tb);
                    mf.ptpSZIforSZI.Controls.Add(gb);

                    i++;
                }
            }
            mf.wsm.Visible = false;
        }

        private void BtnGetSZIs_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            mf.dgvSZIs.Rows.Clear();
            mf.tabControlSZIs.TabPages.Clear();
            mf.tabControlSZIs.TabPages.Add(mf.tpSZItpSZIs);
            int i = 0;
            mf.tpSZItbInfo.Text = "В таблице представлен перечень выбранных Вами средств защиты и краткая информация о них. Для формирования перечня реализуемых технических мер этими средствами нажмите \"Реализация тех. мер\"";
            using (KPSZIContext db = new KPSZIContext())
            {
                var tempListSZIs = db.SZIs.ToList().Intersect(IS.listOfSZIs).ToList();

                foreach (SZI szi in tempListSZIs)
                {
                    string sorts = "";
                    foreach (SZISort ss in szi.SZISorts)
                        sorts += ss.Name + ", ";
                    sorts = sorts.TrimEnd(' ');
                    sorts = sorts.TrimEnd(',');
                    mf.dgvSZIs.Rows.Add(++i, szi.Name, sorts, "№"+szi.Certificate + ", до " + szi.DateOfEnd.ToShortDateString(), szi.NDVControlLevel, szi.SVTClass, szi.TU);
                }
            }
            mf.wsm.Visible = false;
        }

        private void BtnGetMeasSZIs_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            mf.tabControlSZIs.TabPages.Clear();
            mf.tabControlSZIs.TabPages.Add(mf.tpSZItpSZIs);
            mf.tabControlSZIs.TabPages.Add(mf.tpSZItpMeasSZIs);
            mf.tabControlSZIs.SelectedTab = mf.tpSZItpMeasSZIs;
            mf.dgvMeasSZIs.Rows.Clear();
            mf.tpSZItbInfo.Text = "В таблице представлен перечень реализуемых Вами технических мер выбранными Вами средствамизащиты. Для экспорта технического проекта нажмите на кнопку с иконкой Microsoft Office Word";
            using (KPSZIContext db = new KPSZIContext())
            {
                int i = 0;
                var listSZISorts = db.SZISorts.ToList();

                var listMeasues = db.GisMeasures.ToList().Intersect(IS.listOfAllNSDMeasures).ToList();

                foreach (GISMeasure gm in listMeasues)
                {
                    if (gm.SZISorts.Count == 0)
                        continue;

                    string szis = "";
                    List<SZI> _szis = new List<SZI>();
                    foreach (SZISort ss in gm.SZISorts)
                    {
                        _szis = ss.SZIs.ToList().Intersect(IS.listOfSZIs).ToList();

                    }
                    _szis = _szis.Distinct().ToList();

                    foreach (SZI szi in _szis)
                    {
                        szis += szi.Name + "\n";
                    }
                    mf.dgvMeasSZIs.Rows.Add(++i, gm.ToString(), szis);
                }
            }
            mf.wsm.Visible = false;
        }

        private void Rb_CheckedChanged(object sender, EventArgs e)
        {
            foreach (RadioButton rb in radiobuttonsSZISorts)
            {
                SZI szi = listOfSZIsFromDB.Where(s => s.Name == rb.Text).First();
                if (rb.Checked && !IS.listOfSZIs.Contains(szi))
                    IS.listOfSZIs.Add(szi);
                if (!rb.Checked && IS.listOfSZIs.Contains(szi))
                    IS.listOfSZIs.Remove(szi);
            }

            foreach (CheckBox cb in checkboxesSZISorts)
            {
                SZI szi = listOfSZIsFromDB.Where(s => s.Name == cb.Text).First();
                if (cb.Checked && !IS.listOfSZIs.Contains(szi))
                    IS.listOfSZIs.Add(szi);
                if (!cb.Checked && IS.listOfSZIs.Contains(szi))
                    IS.listOfSZIs.Remove(szi);
            }
        }

        private void TabControlSZIs_Selecting(object sender, TabControlCancelEventArgs e)
        {
            //TabPage tp = (TabPage)sender;
            
            if (e.TabPage == mf.tpSZItpReq)
                mf.tpSZItbInfo.Text = "Ниже представлены сформированные требования к средствам защиты информации согласно классу защищенности информационной системы и условиям функционирования. В соответствии с требованиям сформирован перечень видов СЗИ и их представителей. После осуществления выбора, нажмите \"СЗИ\", чтобы просмотреть информацию о них.";

            if (e.TabPage == mf.tpSZItpSZIs)
                mf.tpSZItbInfo.Text = "В таблице представлен перечень выбранных Вами средств защиты и краткая информация о них. Для формирования перечня реализуемых технических мер этими средствами нажмите \"Реализация тех. мер\"";

            if (e.TabPage == mf.tpSZItpMeasSZIs)
                mf.tpSZItbInfo.Text = "В таблице представлен перечень реализуемых технических мер выбранными Вами средствами защиты. Для экспорта технического проекта нажмите на кнопку с иконкой Microsoft Office Word";

        }

        private void BtnExportTP_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            Microsoft.Office.Interop.Word._Application oWord = null;
            try
            {
                oWord = new Microsoft.Office.Interop.Word.Application();
            }
            catch
            {
                mf.wsm.Visible = false;
                MessageBox.Show("На ПК не установлен пакет Microsoft Office Word 2007 или позднее. Экспорт невозможен.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Word._Document oDoc = null;
            try
            {
                oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\templateTP.docx");
            }
            catch
            {
                mf.wsm.Visible = false;
                MessageBox.Show("Отсутствует файл шаблона тех. проекта \"templateTP.docx\". Экспорт невозможен.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Word.Table wordTable;
            Microsoft.Office.Interop.Word.Range bm;
            int row = 0;
            using (KPSZIContext db = new KPSZIContext())
            {

                #region Заполнение СФХ
                var listOfSFHs = db.SFHs.ToList().Intersect(IS.listOfSFHs).ToList();
                var listOfSFHTypes = new List<SFHType>();

                foreach (SFH sfh in listOfSFHs)
                    listOfSFHTypes.Add(sfh.SFHType);
                listOfSFHTypes = listOfSFHTypes.Distinct().OrderBy(st => st.SFHTypeId).ToList();

                bm = oDoc.Bookmarks["SFH"].Range;
                bm.Tables.Add(bm, listOfSFHTypes.Count, 2, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                List<SFHType> lst = new List<SFHType>();

                row = 1;

                foreach (SFHType sfhtype in listOfSFHTypes)
                {
                    if (IS.listOfSFHs.Where(s => s.SFHType.Name == sfhtype.Name).ToList().Count == 0)
                        continue;

                    wordTable.Cell(row, 1).Range.Text = sfhtype.Name;
                    string sfhs = "";
                    foreach (SFH sfh in IS.listOfSFHs.Where(sfh => sfh.SFHType.Name == sfhtype.Name).ToList())
                    {
                        sfhs += sfh.Name + ", ";
                    }
                    sfhs = sfhs.TrimEnd(' ').TrimEnd(',');
                    wordTable.Cell(row, 2).Range.Text = sfhs;
                    row++;
                }

                #endregion

                #region Базовый набор мер
                bm = oDoc.Bookmarks["Basic_Set"].Range;
                bm.Tables.Add(bm, mf.dgvBasicMeas.Rows.Count + 1, 2, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 94f;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№ п/п";
                wordTable.Cell(1, 2).Range.Text = "Наименование базовой меры";

                row = 2;
                foreach (DataGridViewRow dgvrow in mf.dgvBasicMeas.Rows)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row++, 2).Range.Text = dgvrow.Cells[1].Value.ToString();
                }

                #endregion

                #region Адаптация базового набора мер
                bm = oDoc.Bookmarks["Adaptive_Set"].Range;
                bm.Tables.Add(bm, StageMeasures.ListOfExcludedMeasures.Count + 1, 2, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 94f;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№ п/п";
                wordTable.Cell(1, 2).Range.Text = "Наименование меры";

                row = 2;
                foreach (GISMeasure gm in StageMeasures.ListOfExcludedMeasures)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row++, 2).Range.Text = gm.ToString();
                }
                #endregion

                #region Уточненение адаптированного базового набора мер
                //Вывод УБИ-меры
                bm = oDoc.Bookmarks["Thr_Meas"].Range;
                bm.Tables.Add(bm, mf.dgvThrMeas.Rows.Count + 1, 3, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 19f;
                wordTable.Columns[3].PreferredWidth = 75f;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№ п/п";
                wordTable.Cell(1, 2).Range.Text = "Наименование угрозы";
                wordTable.Cell(1, 3).Range.Text = "Меры по нейтрализации УБИ";

                row = 2;
                foreach (DataGridViewRow dgvrow in mf.dgvThrMeas.Rows)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row, 2).Range.Text = dgvrow.Cells[1].Value.ToString();
                    wordTable.Cell(row++, 3).Range.Text = dgvrow.Cells[2].Value.ToString();
                }

                //Вывод добавляемых мер
                bm = oDoc.Bookmarks["Add_Meas"].Range;
                bm.Tables.Add(bm, StageMeasures.ListOfAddedMeasures.Count + 1, 2, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 94f;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№ п/п";
                wordTable.Cell(1, 2).Range.Text = "Наименование добавляемой меры";

                row = 2;
                foreach (GISMeasure gm in StageMeasures.ListOfAddedMeasures)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row++, 2).Range.Text = gm.ToString();
                }
                #endregion

                #region Итоговый перечень мер
                bm = oDoc.Bookmarks["Final_Set"].Range;
                bm.Tables.Add(bm, mf.dgvConcreteMeas.Rows.Count + 1, 2, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 94f;
                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№ п/п";
                wordTable.Cell(1, 2).Range.Text = "Наименование меры";

                row = 2;
                foreach (DataGridViewRow dgvrow in mf.dgvConcreteMeas.Rows)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row++, 2).Range.Text = dgvrow.Cells[1].Value.ToString();
                }
                #endregion

                #region Перечень предлагаемых СЗИ
                bm = oDoc.Bookmarks["SZIs_Set"].Range;
                bm.Tables.Add(bm, mf.dgvSZIs.Rows.Count + 1, 5, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 4.5f;
                wordTable.Columns[2].PreferredWidth = 36.3f;
                wordTable.Columns[3].PreferredWidth = 19.6f;
                wordTable.Columns[4].PreferredWidth = 25.7f;
                wordTable.Columns[5].PreferredWidth = 13.6f;

                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = "№";
                wordTable.Cell(1, 2).Range.Text = "Наименование СЗИ";
                wordTable.Cell(1, 3).Range.Text = "Вид";
                wordTable.Cell(1, 4).Range.Text = "Сертификат";
                wordTable.Cell(1, 5).Range.Text = "Уровень контроля отсутствия НДВ";


                row = 2;
                foreach (DataGridViewRow dgvrow in mf.dgvSZIs.Rows)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row, 2).Range.Text = dgvrow.Cells[1].Value.ToString();
                    wordTable.Cell(row, 3).Range.Text = dgvrow.Cells[2].Value.ToString();
                    wordTable.Cell(row, 4).Range.Text = dgvrow.Cells[3].Value.ToString();
                    wordTable.Cell(row++, 5).Range.Text = dgvrow.Cells[4].Value.ToString();

                }
                #endregion

                #region Мера-СЗИ
                bm = oDoc.Bookmarks["Meas_SZIs"].Range;
                bm.Tables.Add(bm, mf.dgvMeasSZIs.Rows.Count + 1, 3, Type.Missing, Type.Missing);
                wordTable = bm.Tables[1];
                wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                wordTable.Columns[1].PreferredWidth = 6f;
                wordTable.Columns[2].PreferredWidth = 60f;
                wordTable.Columns[3].PreferredWidth = 34f;

                wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
                wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
                wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wordTable.Range.Select();

                wordTable.Cell(1, 1).Range.Text = @"№ п\п";
                wordTable.Cell(1, 2).Range.Text = "Название меры";
                wordTable.Cell(1, 3).Range.Text = "Перечень СЗИ";


                row = 2;
                foreach (DataGridViewRow dgvrow in mf.dgvMeasSZIs.Rows)
                {
                    wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                    wordTable.Cell(row, 2).Range.Text = dgvrow.Cells[1].Value.ToString();
                    wordTable.Cell(row++, 3).Range.Text = dgvrow.Cells[2].Value.ToString();

                }
                #endregion

                #region Замена слов
                mf.FindAndReplace(oWord, "{ИМЯ_ИС}", IS.ISName);
                mf.FindAndReplace(oWord, "{КЗ}", IS.GISClass);

                #endregion


                oDoc.Activate();
                oWord.Visible = true;
            }
            mf.wsm.Visible = false;
        }

        #endregion
    }
}
