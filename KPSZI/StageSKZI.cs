using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KPSZI
{
    public enum SKZIClass
    {
        Неопределен = 0, КС1=1, КС2=2, КС3=3, КВ=4, КА=5
    }

    public class SKZIMeasure
    {
        public int number;
        public string description;

        public static SKZIMeasure[] returnMeasures()
        {
            SKZIMeasure[] measures = new SKZIMeasure[17];
            SKZIMeasure skm1 = new SKZIMeasure { number = 1, description = "проводятся работы по подбору персонала" };
            SKZIMeasure skm2 = new SKZIMeasure { number = 2, description = "доступ в контролируемую зону, где располагается СКЗИ, обеспечивается в соответствии с контрольно-пропускным режимом" };
            SKZIMeasure skm3 = new SKZIMeasure { number = 3, description = "представители технических, обслуживающих и других вспомогательных служб при работе в помещениях (стойках), где расположены СКЗИ, и сотрудники, не являющиеся пользователями СКЗИ, находятся в этих помещениях только в присутствии сотрудников по эксплуатации" };
            SKZIMeasure skm4 = new SKZIMeasure { number = 4, description = "сотрудники, являющиеся пользователями ИС, но не являющиеся пользователями СКЗИ, проинформированы о правилах работы в ИС и ответственности за несоблюдение правил обеспечения безопасности информации" };
            SKZIMeasure skm5 = new SKZIMeasure { number = 5, description = "пользователи СКЗИ проинформированы о правилах работы в ИС, правилах работы с СКЗИ и ответственности за несоблюдение правил обеспечения безопасности информации" };
            SKZIMeasure skm6 = new SKZIMeasure { number = 6, description = "утверждены правила доступа в помещения, где располагаются СКЗИ, в рабочее и нерабочее время, а также в нештатных ситуациях" };
            SKZIMeasure skm7 = new SKZIMeasure { number = 7, description = "утвержден перечень лиц, имеющих право доступа в помещения, где располагаются СКЗИ" };
            SKZIMeasure skm8 = new SKZIMeasure { number = 8, description = "осуществляется разграничение и контроль доступа пользователей к защищаемым ресурсам" };
            SKZIMeasure skm9 = new SKZIMeasure { number = 9, description = "осуществляется регистрация и учет действий пользователей с ПДн" };
            SKZIMeasure skm10 = new SKZIMeasure { number = 10, description = "осуществляется контроль целостности средств защиты на АРМ и серверах, на которых установлены СКЗИ: используются сертифицированные средства защиты информации от несанкционированного доступа, используются сертифицированные средства антивирусной защиты" };
            SKZIMeasure skm11 = new SKZIMeasure { number = 11, description = "документация на СКЗИ хранится у ответственного за СКЗИ в металлическом сейфе" };
            SKZIMeasure skm12 = new SKZIMeasure { number = 12, description = "помещение, в которых располагаются документация на СКЗИ, СКЗИ и компоненты СФ, оснащены входными дверьми с замками, обеспечения постоянного закрытия дверей помещений на замок и их открытия только для санкционированного прохода" };
            SKZIMeasure skm13 = new SKZIMeasure { number = 13, description = "помещения, в которых располагаются СКЗИ, оснащены входными дверьми с замками, обеспечения постоянного закрытия дверей помещений на замок и их открытия только для санкционированного прохода" };
            SKZIMeasure skm14 = new SKZIMeasure { number = 14, description = "сотрудники проинформированы об ответственности за несоблюдение правил обеспечения безопасности" };
            SKZIMeasure skm15 = new SKZIMeasure { number = 15, description = "осуществляется регистрация и учет действий пользователей" };
            SKZIMeasure skm16 = new SKZIMeasure { number = 16, description = "не осуществляется обработка сведений, составляющих государственную тайну, а также иных сведений, которые могут представлять интерес для реализации возможности" };
            SKZIMeasure skm17 = new SKZIMeasure { number = 17, description = "высокая стоимость и сложность подготовки реализации возможности" };
            measures[0] = skm1;
            measures[1] = skm2;
            measures[2] = skm3;
            measures[3] = skm4;
            measures[4] = skm5;
            measures[5] = skm6;
            measures[6] = skm7;
            measures[7] = skm8;
            measures[8] = skm9;
            measures[9] = skm10;
            measures[10] = skm11;
            measures[11] = skm12;
            measures[12] = skm13;
            measures[13] = skm14;
            measures[14] = skm15;
            measures[15] = skm16;
            measures[16] = skm17;
            return measures;
        }
    }

    class myNormalPanel : Panel
    {
        protected override System.Drawing.Point ScrollToControl(Control activeControl)
        {
            return this.AutoScrollPosition;
        }
    }

    class StageSKZI : Stage
    {
        private Form dialogFormSKZI;

        private int currentMaxCheckedAbility;

        public StageSKZI(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS) : base(stageTab, stageNode, mainForm, IS)
        {

        }

        protected override ImageList imageListForTabPage { get; set; }

        public override void enterTabPage()
        {
            mf.tcSKZI.SelectedTab = mf.tcSKZI.TabPages[0];
        }

        public override void saveChanges()
        {

        }

        public void initDialofFormSKZIMeasures()
        {
            SKZIMeasure[] measures = SKZIMeasure.returnMeasures();

            dialogFormSKZI = new Form();
            dialogFormSKZI.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialogFormSKZI.Icon = KPSZI.Properties.Resources.mf;
            dialogFormSKZI.MaximizeBox = false;
            dialogFormSKZI.MinimizeBox = false;

            CheckBox checkBox1 = new CheckBox();
            CheckBox checkBox2 = new CheckBox();
            CheckBox checkBox3 = new CheckBox();
            CheckBox checkBox4 = new CheckBox();
            CheckBox checkBox5 = new CheckBox();
            CheckBox checkBox6 = new CheckBox();
            CheckBox checkBox7 = new CheckBox();
            CheckBox checkBox8 = new CheckBox();
            CheckBox checkBox9 = new CheckBox();
            CheckBox checkBox10 = new CheckBox();
            CheckBox checkBox11 = new CheckBox();
            CheckBox checkBox12 = new CheckBox();
            CheckBox checkBox13 = new CheckBox();
            CheckBox checkBox14 = new CheckBox();
            CheckBox checkBox15 = new CheckBox();
            CheckBox checkBox16 = new CheckBox();
            CheckBox checkBox17 = new CheckBox();
            Button button1 = new Button();
            Label label1 = new Label();
            Label RowIndex = new Label();
            RowIndex.Visible = false;
            RowIndex.Name = "RowIndex";
            Label ColumnIndex = new Label();
            ColumnIndex.Visible = false;
            ColumnIndex.Name = "ColumnIndex";
            // label1
            // 
            label1.Location = new System.Drawing.Point(13, 470);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(663, 75);
            label1.TabIndex = 18;
            label1.Font = new System.Drawing.Font(label1.Font.FontFamily, 9);
            label1.Text = "";
            //
            // button1
            //
            button1.Click += new System.EventHandler(acceptButtonClick);
            button1.Location = new System.Drawing.Point(752, 470);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(75, 23);
            button1.TabIndex = 17;
            button1.Text = "Принять";
            button1.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            checkBox1.Location = new System.Drawing.Point(12, 12);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new System.Drawing.Size(250, 24);
            checkBox1.TabIndex = 0;
            checkBox1.Text = measures[0].description;
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            checkBox2.Location = new System.Drawing.Point(12, 46);
            checkBox2.Name = "checkBox2";
            checkBox2.Size = new System.Drawing.Size(263, 63);
            checkBox2.TabIndex = 1;
            checkBox2.Text = measures[1].description;
            checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            checkBox3.Location = new System.Drawing.Point(12, 115);
            checkBox3.Name = "checkBox3";
            checkBox3.Size = new System.Drawing.Size(263, 110);
            checkBox3.TabIndex = 2;
            checkBox3.Text = measures[2].description;
            checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            checkBox4.Location = new System.Drawing.Point(12, 231);
            checkBox4.Name = "checkBox4";
            checkBox4.Size = new System.Drawing.Size(263, 86);
            checkBox4.TabIndex = 3;
            checkBox4.Text = measures[3].description;
            checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            checkBox5.Location = new System.Drawing.Point(12, 323);
            checkBox5.Name = "checkBox5";
            checkBox5.Size = new System.Drawing.Size(263, 73);
            checkBox5.TabIndex = 4;
            checkBox5.Text = measures[4].description;
            checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            checkBox6.Location = new System.Drawing.Point(12, 402);
            checkBox6.Name = "checkBox6";
            checkBox6.Size = new System.Drawing.Size(250, 65);
            checkBox6.TabIndex = 5;
            checkBox6.Text = measures[5].description;
            checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox7
            // 
            checkBox7.Location = new System.Drawing.Point(330, 12);
            checkBox7.Name = "checkBox7";
            checkBox7.Size = new System.Drawing.Size(255, 53);
            checkBox7.TabIndex = 6;
            checkBox7.Text = measures[6].description;
            checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox8
            // 
            checkBox8.Location = new System.Drawing.Point(330, 71);
            checkBox8.Name = "checkBox8";
            checkBox8.Size = new System.Drawing.Size(255, 51);
            checkBox8.TabIndex = 7;
            checkBox8.Text = measures[7].description;
            checkBox8.UseVisualStyleBackColor = true;
            // 
            // checkBox9
            // 
            checkBox9.Location = new System.Drawing.Point(627, 284);
            checkBox9.Name = "checkBox9";
            checkBox9.Size = new System.Drawing.Size(200, 50);
            checkBox9.TabIndex = 8;
            checkBox9.Text = measures[8].description;
            checkBox9.UseVisualStyleBackColor = true;
            // 
            // checkBox10
            // 
            checkBox10.Location = new System.Drawing.Point(330, 128);
            checkBox10.Name = "checkBox10";
            checkBox10.Size = new System.Drawing.Size(255, 114);
            checkBox10.TabIndex = 9;
            checkBox10.Text = measures[9].description;
            checkBox10.UseVisualStyleBackColor = true;
            // 
            // checkBox11
            // 
            checkBox11.Location = new System.Drawing.Point(627, 340);
            checkBox11.Name = "checkBox11";
            checkBox11.Size = new System.Drawing.Size(200, 49);
            checkBox11.TabIndex = 10;
            checkBox11.Text = measures[10].description;
            checkBox11.UseVisualStyleBackColor = true;
            // 
            // checkBox12
            // 
            checkBox12.Location = new System.Drawing.Point(330, 231);
            checkBox12.Name = "checkBox12";
            checkBox12.Size = new System.Drawing.Size(255, 113);
            checkBox12.TabIndex = 11;
            checkBox12.Text = measures[11].description;
            checkBox12.UseVisualStyleBackColor = true;
            // 
            // checkBox13
            // 
            checkBox13.Location = new System.Drawing.Point(330, 350);
            checkBox13.Name = "checkBox13";
            checkBox13.Size = new System.Drawing.Size(255, 91);
            checkBox13.TabIndex = 12;
            checkBox13.Text = measures[12].description;
            checkBox13.UseVisualStyleBackColor = true;
            // 
            // checkBox14
            // 
            checkBox14.Location = new System.Drawing.Point(627, 12);
            checkBox14.Name = "checkBox14";
            checkBox14.Size = new System.Drawing.Size(200, 58);
            checkBox14.TabIndex = 13;
            checkBox14.Text = measures[13].description;
            checkBox14.UseVisualStyleBackColor = true;
            // 
            // checkBox15
            // 
            checkBox15.Location = new System.Drawing.Point(627, 85);
            checkBox15.Name = "checkBox15";
            checkBox15.Size = new System.Drawing.Size(188, 39);
            checkBox15.TabIndex = 14;
            checkBox15.Text = measures[14].description;
            checkBox15.UseVisualStyleBackColor = true;
            // 
            // checkBox16
            // 
            checkBox16.Location = new System.Drawing.Point(627, 130);
            checkBox16.Name = "checkBox16";
            checkBox16.Size = new System.Drawing.Size(188, 88);
            checkBox16.TabIndex = 15;
            checkBox16.Text = measures[15].description;
            checkBox16.UseVisualStyleBackColor = true;
            // 
            // checkBox17
            // 
            checkBox17.Location = new System.Drawing.Point(627, 224);
            checkBox17.Name = "checkBox17";
            checkBox17.Size = new System.Drawing.Size(188, 54);
            checkBox17.TabIndex = 16;
            checkBox17.Text = measures[16].description;
            checkBox17.UseVisualStyleBackColor = true;
            // 
            // dialogFormSKZI
            // 
            dialogFormSKZI.AutoScroll = true;
            dialogFormSKZI.ClientSize = new System.Drawing.Size(856, 550);
            dialogFormSKZI.Controls.Add(checkBox17);
            dialogFormSKZI.Controls.Add(checkBox16);
            dialogFormSKZI.Controls.Add(checkBox15);
            dialogFormSKZI.Controls.Add(checkBox14);
            dialogFormSKZI.Controls.Add(checkBox13);
            dialogFormSKZI.Controls.Add(checkBox12);
            dialogFormSKZI.Controls.Add(checkBox11);
            dialogFormSKZI.Controls.Add(checkBox10);
            dialogFormSKZI.Controls.Add(checkBox9);
            dialogFormSKZI.Controls.Add(checkBox8);
            dialogFormSKZI.Controls.Add(checkBox7);
            dialogFormSKZI.Controls.Add(checkBox6);
            dialogFormSKZI.Controls.Add(checkBox5);
            dialogFormSKZI.Controls.Add(checkBox4);
            dialogFormSKZI.Controls.Add(checkBox3);
            dialogFormSKZI.Controls.Add(checkBox2);
            dialogFormSKZI.Controls.Add(checkBox1);
            dialogFormSKZI.Controls.Add(button1);
            dialogFormSKZI.Controls.Add(label1);
            dialogFormSKZI.Controls.Add(RowIndex);
            dialogFormSKZI.Controls.Add(ColumnIndex);
            dialogFormSKZI.Name = "dialogFormSKZI";
            dialogFormSKZI.Text = "Обоснование неактуальности угроз";
        }

        protected override void initTabPage()
        {
            mf.dgvSKZIUtochnAbils.CellValueChanged += new DataGridViewCellEventHandler(cellValueChanged);
            myNormalPanel mnPanelTab2 = new myNormalPanel();
            mnPanelTab2.Dock = DockStyle.Fill;
            mnPanelTab2.Location = new System.Drawing.Point(0, 0);
            mnPanelTab2.AutoScroll = true;
            mnPanelTab2.Controls.Add(mf.dgvSKZIAttackAbils);
            mf.tcSKZI.TabPages[1].Controls.Add(mnPanelTab2);

            myNormalPanel mnPanelTab3 = new myNormalPanel();
            mnPanelTab3.Dock = DockStyle.Fill;
            mnPanelTab3.Location = new System.Drawing.Point(0, 0);
            mnPanelTab3.AutoScroll = true;
            mnPanelTab3.Controls.Add(mf.dgvSKZIUtochnAbils);
            mf.tcSKZI.TabPages[2].Controls.Add(mnPanelTab3);

            #region Строки 2 вкладки
            DataGridViewRow dgr21 = new DataGridViewRow();
            dgr21.CreateCells(mf.dgvSKZIAttackAbils);
            dgr21.Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgr21.Cells[0].Value = "1";
            dgr21.Cells[1].Value = "Возможность самостоятельно осуществлять создание способов атак, подготовку и проведение атак только за пределами контролируемой зоны";
            dgr21.Cells[3].Value = "Нарушитель с высоким потенциалом";

            DataGridViewRow dgr22 = new DataGridViewRow();
            dgr22.CreateCells(mf.dgvSKZIAttackAbils);
            dgr22.Cells[0].Value = "2";
            dgr22.Cells[1].Value = "Возможность самостоятельно осуществлять создание способов атак, подготовку и проведение атак в пределах контролируемой зоны, но без физического доступа к аппаратным средствам(далее – АС), на которых реализованы СКЗИ и среда их функционирования";
            dgr22.Cells[3].Value = "Нарушитель с высоким потенциалом";

            DataGridViewRow dgr23 = new DataGridViewRow();
            dgr23.CreateCells(mf.dgvSKZIAttackAbils);
            dgr23.Cells[0].Value = "3";
            dgr23.Cells[1].Value = "Возможность самостоятельно осуществлять создание способов атак, подготовку и проведение атак в пределах контролируемой зоны с физическим доступом к АС, на которых реализованы СКЗИ и среда их функционирования";
            dgr23.Cells[2].Value = true;
            dgr23.Cells[3].Value = "Нарушитель с базовым (низким) потенциалом и нарушители с базовым повышенным (средним) потенциалом";

            DataGridViewRow dgr24 = new DataGridViewRow();
            dgr24.CreateCells(mf.dgvSKZIAttackAbils);
            dgr24.Cells[0].Value = "4";
            dgr24.Cells[1].Value = "Возможность привлекать специалистов, имеющих опыт разработки и анализа СКЗИ (включая специалистов в области анализа сигналов линейной передачи и сигналов побочного электромагнитного излучения и наводок СКЗИ)";
            dgr24.Cells[3].Value = "Нарушитель с высоким потенциалом";

            DataGridViewRow dgr25 = new DataGridViewRow();
            dgr25.CreateCells(mf.dgvSKZIAttackAbils);
            dgr25.Cells[0].Value = "5";
            dgr25.Cells[1].Value = "Возможность привлекать специалистов, имеющих опыт разработки и анализа СКЗИ (включая специалистов в области использования для реализации атак недокументированных возможностей прикладного программного обеспечения)";
            dgr25.Cells[3].Value = "Нарушитель с высоким потенциалом";

            DataGridViewRow dgr26 = new DataGridViewRow();
            dgr26.CreateCells(mf.dgvSKZIAttackAbils);
            dgr26.Cells[0].Value = "6";
            dgr26.Cells[1].Value = "Возможность привлекать специалистов, имеющих опыт разработки и анализа СКЗИ (включая специалистов в области использования для реализации атак недокументированных возможностей аппаратного и программного компонентов среды функционирования СКЗИ)";
            dgr26.Cells[3].Value = "Нарушитель с высоким потенциалом";
            #endregion

            
            mf.dgvSKZIAttackAbils.Rows.AddRange(new DataGridViewRow[] { dgr21, dgr22, dgr23, dgr24, dgr25, dgr26 });

            #region Строки 3 вкладки
            DataGridViewRow dgr1 = new DataGridViewRow();
            dgr1.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr1.Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgr1.Cells[0].Value = "1.1";
            dgr1.Cells[1].Value = "проведение атаки при нахождении в пределах контролируемой зоны";
            dgr1.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr1.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr2 = new DataGridViewRow();
            dgr2.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr2.Cells[0].Value = "1.2";
            dgr2.Cells[1].Value = "проведение атак на этапе эксплуатации СКЗИ на следующие объекты: \n\tдокументацию на СКЗИ и компоненты СФ;\n\t-помещения, в которых находится совокупность программных и технических элементов систем обработки данных, способных функционировать самостоятельно или в составе других систем(далее - СВТ), на которых реализованы СКЗИ и СФ;";
            dgr2.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr2.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr3 = new DataGridViewRow();
            dgr3.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr3.Cells[0].Value = "1.3";
            dgr3.Cells[1].Value = "получение в рамках предоставленных полномочий, а также в результате наблюдений следующей информации: \n\t- сведений о физических мерах защиты объектов, в которых размещены ресурсы информационной системы; \n\t- сведений о мерах по обеспечению контролируемой зоны объектов, в которых размещены ресурсы информационной системы; \n\t- сведений о мерах по разграничению доступа в помещения,в которых находятся СВТ, накоторых реализованы СКЗИ и СФ;";
            dgr3.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr3.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr4 = new DataGridViewRow();
            dgr4.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr4.Cells[0].Value = "1.4";
            dgr4.Cells[1].Value = "использование штатных средств ИСПДн, ограниченное мерами, реализованными в информационной системе, в которой используется СКЗИ, и направленными на предотвращение и  пресечение несанкционированных действий.";
            dgr4.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr4.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr5 = new DataGridViewRow();
            dgr5.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr5.Cells[0].Value = "2.1";
            dgr5.Cells[1].Value = "физический доступ к СВТ, на которых реализованы СКЗИ и СФ; ";
            dgr5.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr5.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr6 = new DataGridViewRow();
            dgr6.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr6.Cells[0].Value = "2.2";
            dgr6.Cells[1].Value = "возможность воздействовать на аппаратные компоненты СКЗИ и СФ, ограниченная мерами, реализованными в информационной системе, в которой используется СКЗИ, и направленными на предотвращение и пресечение несанкционированных действий;";
            dgr6.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr6.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr7 = new DataGridViewRow();
            dgr7.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr7.Cells[0].Value = "2.2";
            dgr7.Cells[1].Value = "возможность воздействовать на аппаратные компоненты СКЗИ и СФ,ограниченная мерами, реализованными в информационной системе, в которой используется СКЗИ, и направленными на предотвращение и пресечение несанкционированных действий";
            dgr7.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr7.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr8 = new DataGridViewRow();
            dgr8.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr8.Cells[0].Value = "3.1";
            dgr8.Cells[1].Value = "создание способов, подготовка и проведение атак с привлечением специалистов в области анализа сигналов, сопровождающих функционирование СКЗИ и СФ, и в области спользования для реализации атак недокументированных (недекларированных) возможностей прикладного ПО;";
            dgr8.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr8.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr9 = new DataGridViewRow();
            dgr9.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr9.Cells[0].Value = "3.2";
            dgr9.Cells[1].Value = "проведение лабораторных исследований СКЗИ, используемых вне контролируемой зоны, ограниченное мерами, реализованными в информационной системе, в которой используется СКЗИ, и направленными на предотвращение и пресечение несанкционированных действий;";
            dgr9.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr9.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr10 = new DataGridViewRow();
            dgr10.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr10.Cells[0].Value = "3.3";
            dgr10.Cells[1].Value = "проведение работ по созданию способов и средств атак в научноисследовательских центрах, специализирующихся в области разработки и анализа СКЗИ и СФ, в том числе с использованием исходных текстов входящего в СФ прикладного ПО, непосредственно использующего вызовы программных функций СКЗИ;";
            dgr10.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr10.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr11 = new DataGridViewRow();
            dgr11.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr11.Cells[0].Value = "4.1";
            dgr11.Cells[1].Value = "создание способов, подготовка и проведение атак с привлечением специалистов в области использования для реализации атак недокументированных (недекларированных) возможностей системного ПО;";
            dgr11.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr11.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr12 = new DataGridViewRow();
            dgr12.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr12.Cells[0].Value = "4.2";
            dgr12.Cells[1].Value = "возможность располагать сведениями, содержащимися в конструкторской документации на аппаратные и программные компоненты СФ;";
            dgr12.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr12.Cells[2])).Value = "Не актуально";

            DataGridViewRow dgr13 = new DataGridViewRow();
            dgr13.CreateCells(mf.dgvSKZIUtochnAbils);
            dgr13.Cells[0].Value = "4.3";
            dgr13.Cells[1].Value = "возможность воздействовать на любые компоненты СКЗИ и СФ";
            dgr13.Cells[3].Value = "";
            ((DataGridViewComboBoxCell)(dgr13.Cells[2])).Value = "Не актуально";
            #endregion
            mf.dgvSKZIUtochnAbils.Rows.AddRange(new DataGridViewRow[] { dgr1, dgr2, dgr3, dgr4, dgr5, dgr6, dgr7, dgr8, dgr9, dgr10, dgr11, dgr12, dgr13 });
            mf.dgvSKZIUtochnAbils.EditMode = DataGridViewEditMode.EditOnEnter;
            mf.tcSKZI.TabPages[1].Enter += new EventHandler(enterTab2);
            mf.tcSKZI.TabPages[2].Enter += new EventHandler(enterTab3);

            mf.dgvSKZIUtochnAbils.CellClick += new DataGridViewCellEventHandler(dgvUtochAbilsCellClick);

            mf.dgvSKZIAttackAbils.CellValueChanged += new DataGridViewCellEventHandler(changedMaxCheckedAbility);
            changedMaxCheckedAbility(null, new DataGridViewCellEventArgs(2, 3));
            SetHeightOfDGV(mf.dgvSKZIUtochnAbils);
        }
        
        public void cellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
                calculateThreats();
        }

        public void acceptButtonClick(object sender, EventArgs e)
        {
            string measures = "";
            foreach (Control cb in dialogFormSKZI.Controls)
            {
                if (cb as CheckBox != null)
                {
                    if (((CheckBox)cb).Checked)
                        measures += cb.Text + "; ";
                }
            }
            int rowIndex;
            int columnIndex;
            int.TryParse(dialogFormSKZI.Controls["RowIndex"].Text, out rowIndex);
            int.TryParse(dialogFormSKZI.Controls["ColumnIndex"].Text, out columnIndex);
            if (measures != "")
                mf.dgvSKZIUtochnAbils.Rows[rowIndex].Cells[columnIndex].Value = measures;
            else
                mf.dgvSKZIUtochnAbils.Rows[rowIndex].Cells[columnIndex].Value = "Обоснуйте неактуальность уточненной возможности";
            dialogFormSKZI.Close();
            dialogFormSKZI.Dispose();
        }

        public void dgvUtochAbilsCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3 && mf.dgvSKZIUtochnAbils.Rows[e.RowIndex].Cells[3].Value.ToString() != "")
            {
                openSKZIDialogForm(e.RowIndex, e.ColumnIndex);
                SetHeightOfDGV(mf.dgvSKZIUtochnAbils);
                calculateThreats();
            }
            if (e.ColumnIndex == 2)
                calculateThreats();
        }
        
        public void calculateThreats()
        {
            List<string> actualThreatNumbers = new List<string>();
            int maxPossibleActualThreatRow = 0;
            int countActualThreats = 1;
            switch (getHighestCheckedAbility())
            {
                case 3:
                    {
                        maxPossibleActualThreatRow = 5;
                        break;
                    }
                case 4:
                case 5:
                    {
                        maxPossibleActualThreatRow = 9;
                        break;
                    }
                case 6:
                    {
                        maxPossibleActualThreatRow = 12;
                        countActualThreats = 13;
                        break;
                    }
            }

            int actualThreatsInGrid = 0;
            int countNoReasonedThreats = 0;
            foreach (DataGridViewRow row in mf.dgvSKZIUtochnAbils.Rows)
            {
                if (row.Index <= maxPossibleActualThreatRow)
                {
                    if (row.Cells[2].Value.ToString() == "Актуально")
                    {
                        actualThreatNumbers.Add(row.Cells[0].Value.ToString());
                        actualThreatsInGrid++;
                    }
                }
                else
                {
                    if (row.Cells[3].Value.ToString() == "Обоснуйте неактуальность уточненной возможности")
                        countNoReasonedThreats++;
                }
            }

            mf.lbCountSKZIThreats.Text = "Минимальное количество актуальных угроз: " + countActualThreats.ToString() +
                "\r\nОтмечено актуальных угроз: " + actualThreatsInGrid.ToString();
            if (countNoReasonedThreats > 0)
                mf.lbCountSKZIThreats.Text += "\r\nОбоснуйте неактуальность следующего количества угроз " + countNoReasonedThreats;
            if (countNoReasonedThreats == 0 && actualThreatsInGrid >= countActualThreats)
                mf.lbCountSKZIThreats.ForeColor = System.Drawing.Color.Green;
            else
                mf.lbCountSKZIThreats.ForeColor = System.Drawing.Color.Red;

            List<SKZIClass> SKZIClasses = new List<SKZIClass>();

            if ((IS.typeOfActualThreats >= 1 && IS.typeOfActualThreats <= 3) && (IS.levelOfPDN >= 1 && IS.levelOfPDN <= 4))
                foreach (string actualThreat in actualThreatNumbers)
                {
                    SKZIClasses.Add(calculateSKZIClass(actualThreat));
                }
            else
            {
                MessageBox.Show("Определите уровень защищенности персональных данных.");
                mf.treeView.SelectedNode = mf.returnTreeNode("tnClassification");
                return;
            }

            mf.lbSKZIClass.Text = "Класс СКЗИ: " + getMax(SKZIClasses);
        }

        public SKZIClass getMax(List<SKZIClass> list)
        {
            SKZIClass max = SKZIClass.Неопределен;
            foreach(SKZIClass item in list)
            {
                if (item > max)
                    max = item;
            }
            return max;
        }

        public SKZIClass calculateSKZIClass(string actualThreat)
        {
            #region Большой свитч по возврату класса СКЗИ
            switch(actualThreat)
            {
                case "1.1":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС2;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "1.2":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС2;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "1.3":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС2;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "1.4":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС2;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "2.1":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС3;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "2.2":
                    {
                        if (IS.levelOfPDN == 4 || ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2) && IS.typeOfActualThreats == 3))
                            return SKZIClass.КС3;
                        else
                        {
                            if ((IS.levelOfPDN == 3 || IS.levelOfPDN == 2 || IS.levelOfPDN == 1) && IS.typeOfActualThreats == 2)
                                return SKZIClass.КВ;
                            else
                            {
                                if (IS.typeOfActualThreats == 1 && (IS.levelOfPDN == 2 || IS.levelOfPDN == 1))
                                    return SKZIClass.КА;
                                else
                                    return SKZIClass.Неопределен;
                            }
                        }
                    }
                case "3.1":
                    {
                        if (IS.levelOfPDN <= 2 && IS.typeOfActualThreats == 1)
                            return SKZIClass.КА;
                        else
                            return SKZIClass.КВ;
                    }
                case "3.2":
                    {
                        if (IS.levelOfPDN <= 2 && IS.typeOfActualThreats == 1)
                            return SKZIClass.КА;
                        else
                            return SKZIClass.КВ;
                    }
                case "3.3":
                    {
                        if (IS.levelOfPDN <= 2 && IS.typeOfActualThreats == 1)
                            return SKZIClass.КА;
                        else
                            return SKZIClass.КВ;
                    }
                case "4.1":
                    {
                        return SKZIClass.КА;
                    }
                case "4.2":
                    {
                        return SKZIClass.КА;
                    }
                case "4.3":
                    {
                        return SKZIClass.КА;
                    }
            }
            #endregion
            return SKZIClass.Неопределен;
        }

        public void openSKZIDialogForm(int rowIndex, int columnIndex)
        {
            initDialofFormSKZIMeasures();
            string threat = mf.dgvSKZIUtochnAbils.Rows[rowIndex].Cells[1].Value.ToString();
            dialogFormSKZI.Controls["label1"].Text = "Угроза: " + threat;
            dialogFormSKZI.Controls["RowIndex"].Text = rowIndex.ToString();
            dialogFormSKZI.Controls["ColumnIndex"].Text = columnIndex.ToString();
            getCheckedMeasures(rowIndex, columnIndex);
            dialogFormSKZI.ShowDialog();
        }

        public void getCheckedMeasures(int rowIndex, int columnIndex)
        {
            string measures = mf.dgvSKZIUtochnAbils.Rows[rowIndex].Cells[columnIndex].Value.ToString();
            string[] arrayMeasures =  measures.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
            SKZIMeasure[] SKZIms = SKZIMeasure.returnMeasures();
            foreach(string comparable in arrayMeasures)
            {
                for(int i = 0;i < SKZIms.Length; i++)
                {
                    if((comparable)==SKZIms[i].description)
                    {
                        ((CheckBox)(dialogFormSKZI.Controls["checkBox" + (i + 1).ToString()])).Checked = true;
                    }
                }
            }
        }

        public void enterTab2(object sender, EventArgs e)
        {
            int intruderPotential = getMaxIntrPot(IS.listOfSources);
            if (IS.listOfInfoTypes.FindIndex(t=> t.TypeName=="Персональные данные")==-1 || (mf.cbSKZIHran.Checked || mf.cbSKZIPered.Checked) == false)
            {
                MessageBox.Show("Для информационной системы не актуально использование средств криптографической защиты");
                mf.tcSKZI.SelectedTab = mf.tcSKZI.TabPages[0];
                return;
            }

            if (intruderPotential == -1)
            {
                MessageBox.Show("Перед определением актуальных угроз утечки информации по техническим каналам выберите нарушителя в соответствующей вкладке.");
                mf.tcSKZI.SelectedTab = mf.tcSKZI.TabPages[0];
                mf.treeView.SelectedNode = mf.returnTreeNode("tnIntruder");
                return;
            }

            foreach (DataGridViewRow row in mf.dgvSKZIAttackAbils.Rows)
            {
                if (row.Index != 2)
                    row.Cells[2].ReadOnly = intruderPotential >= 2 ? false : true;
                else
                    row.Cells[2].Value = true;
            }

            SetHeightOfDGV(mf.dgvSKZIAttackAbils);
            mf.tcSKZI.TabPages[1].AutoScrollOffset = new System.Drawing.Point(0, 0);
        }
        
        public void enterTab3(object sender, EventArgs e)
        {
            int intruderPotential = getMaxIntrPot(IS.listOfSources);
            if (IS.listOfInfoTypes.FindIndex(t => t.TypeName == "Персональные данные") == -1 || (mf.cbSKZIHran.Checked || mf.cbSKZIPered.Checked) == false)
            {
                MessageBox.Show("Для информационной системы не актуально использование средств криптографической защиты");
                mf.tcSKZI.SelectedTab = mf.tcSKZI.TabPages[0];
                return;
            }

            if (currentMaxCheckedAbility != getHighestCheckedAbility())
            {
                
            }

            currentMaxCheckedAbility = getHighestCheckedAbility();
            SetHeightOfDGV(mf.dgvSKZIUtochnAbils);
            calculateThreats();
        }
        
        public void changedMaxCheckedAbility(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                switch (getHighestCheckedAbility())
                {
                    case 3:
                        {
                            foreach (DataGridViewRow row in mf.dgvSKZIUtochnAbils.Rows)
                            {
                                if (row.Index >= 6)
                                {
                                    row.Cells[3].Value = "Обоснуйте неактуальность уточненной возможности";
                                    row.Cells[2].ReadOnly = true;
                                }
                                else
                                {
                                    row.Cells[2].ReadOnly = false;
                                }
                                row.Cells[2].Value = "Не актуально";
                            }
                            break;
                        }
                    case 4:
                    case 5:
                        {
                            foreach (DataGridViewRow row in mf.dgvSKZIUtochnAbils.Rows)
                            {
                                if (row.Index >= 10)
                                {
                                    row.Cells[3].Value = "Обоснуйте неактуальность уточненной возможности";
                                    row.Cells[2].ReadOnly = true;
                                }
                                else
                                {
                                    row.Cells[3].Value = "";
                                    row.Cells[2].ReadOnly = false;
                                }
                                row.Cells[2].Value = "Не актуально";
                            }
                            break;
                        }
                    case 6:
                        {
                            foreach (DataGridViewRow row in mf.dgvSKZIUtochnAbils.Rows)
                            {
                                row.Cells[2].Value = "Актуально";
                                row.Cells[2].ReadOnly = true;
                                row.Cells[3].Value = "";
                            }
                            break;
                        }
                }
            }
        }

        public int getHighestCheckedAbility()
        {
            int max = -1;
            foreach (DataGridViewRow row in mf.dgvSKZIAttackAbils.Rows)
            {
                if (row.Cells[2].Value != null)
                    if ((bool)(row.Cells[2].Value) == true)
                    {
                        max = 1 + row.Index;
                    }
            }
            return max;
        }

        public int getMaxIntrPot(List<KPSZI.Model.ThreatSource> list)
        {
            int max = -1;
            if (list.Count != 0)
            {
                foreach (KPSZI.Model.ThreatSource ts in list)
                {
                    if (ts.Potencial > max && ts.Potencial != 3)
                        max = ts.Potencial;
                }
                return max;
            }
            else
                return -1;
        }

        public void SetHeightOfDGV(DataGridView Table)
        {
            foreach (DataGridViewRow dgvr in Table.Rows)
            {
                int index = dgvr.Index;
                int d = Table.Rows[index].GetPreferredHeight(index, DataGridViewAutoSizeRowMode.AllCells, true);
                dgvr.Height = d;
            }
            Table.Height = Table.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + Table.ColumnHeadersHeight;
        }
    }
}
