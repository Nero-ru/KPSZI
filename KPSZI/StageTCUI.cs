using System.Windows.Forms;
using KPSZI.Model;
using System.Collections.Generic;
using System;
using System.Linq;
using System.Data;
using System.Drawing;

namespace KPSZI
{
    public enum intruderPotencial { Низкий = 0, Средний = 1, Высокий = 2, Невозможен = 3 };
    

    public class checkedTCUI
    {
        public TCUI tc;
        public int counter;
        
        public checkedTCUI(TCUI TC)
        {
            tc = TC;
            counter = 0;
        }
    }

    class StageTCUI : Stage
    {
        public List<IntruderAbilityControl> controlsIAC;
        public List<TCUIThreat> listOfTCUIThreats;
        protected override ImageList imageListForTabPage { get; set; }
        internal List<checkedTCUI> checkedTCUI;
        Dictionary<string, int[,,]> damageDegreeInput;
        DamageDegreeControl DDControl;

        public StageTCUI(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {

        }

        public override void enterTabPage()
        {
            if (mf.tabControlTCUI.SelectedIndex == 2)
                enterTabPageThreatsList(mf.tabControlTCUI.TabPages[2], null);
        }

        public override void saveChanges()
        {

        }

        protected override void initTabPage()
        {
            mf.tabControlTCUI.TabPages[2].AutoScroll = true;
            mf.dgvActualTCUIThreats.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            controlsIAC = new List<IntruderAbilityControl>();
            listOfTCUIThreats = new List<TCUIThreat>();
            checkedTCUI = new List<KPSZI.checkedTCUI>();

            using (KPSZIContext db = new KPSZIContext())
            {
                foreach (TCUI TC in db.TCUIs.ToList())
                {
                    checkedTCUI ctc = new KPSZI.checkedTCUI(new TCUI { Name = TC.Name, TCUIType = TC.TCUIType, TCUIThreats = TC.TCUIThreats.ToList()});
                    checkedTCUI.Add(ctc);
                }
            } 

            foreach (Control cb in mf.tabControlTCUI.TabPages["tabPageTCUIExist"].Controls)
            {
                if (cb as CheckBox != null)
                    ((CheckBox)cb).Click += new EventHandler(cbClick);
            }
            mf.tabControlTCUI.TabPages["tabPageIntrAbil"].Enter += new System.EventHandler(enterAtPageAbilsOfIntruder);
            mf.tabControlTCUI.TabPages["tabPageListOfTCUIThreats"].Enter += new System.EventHandler(enterTabPageThreatsList);
            mf.tabControlTCUI.TabPages["tabPageIntrAbil"].AutoScroll = true;

            // Добавление контрола определения степени ущерба
            //damageDegreeInput = new Dictionary<string, int[,,]>();
            //foreach (IntruderAbilityControl iac in controlsIAC)
            //    damageDegreeInput.Add(iac.threatName.Substring(0, 6), new int[IS.listOfInfoTypes.Count, 3, 7]);

            //DDControl = new DamageDegreeControl(IS.listOfInfoTypes, listFilteredThreats, damageDegreeInput, mf);
            //DDControl.Location = new Point(mf.tpThreatsNSD2.Width - DDControl.Width, 0);
            //DDControl.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            //mf.tpThreatsNSD2.Controls.Add(DDControl);
        }

        public void cbClick(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            #region ОГРОМНЫЙ switch для добавления/удаления каналов утечки по клику соотв. чекбокса
            switch (cb.Name)
            {
                case "cbApparZakladSyom":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbATS":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                        }
                        break;
                    }
                case "cbBytTech":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbClock":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbCommunOnePlace":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Индукционные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Индукционные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                        }
                        break;
                    }
                case "cbEMIHighFreq":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbEMILowFreq":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbEMINavod":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbKabeliKZ":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbLinesSvyaz":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbLumenSvet":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbMicroPotolok":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbNablyud":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Видовые" && t.tc.TCUIType.Name == "Каналы утечки видовой информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Видовые" && t.tc.TCUIType.Name == "Каналы утечки видовой информации").counter--;
                        }
                        break;
                    }
                case "cbOhrPozh":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbOtherBuilding":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Оптико-электронные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Оптико-электронные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbPEMINorm":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbPhoneOutside":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbProsachSignVOTSS":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbProsachZazem":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        break;
                    }
                case "cbRadio":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbRadioZakladki":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbScheli":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbTransformer":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                        }
                        break;
                    }
                case "cbTruby":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbTSPI":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbVoiceDefense":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Индукционные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Индукционные" && t.tc.TCUIType.Name == "Каналы перехвата информации при ее передаче по каналам связи").counter--;
                        }
                        break;
                    }
                case "cbVoiceSvyaz":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbVTSS":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Электромагнитные" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Параметрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Акусто-электрические" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbWindows":
                    {
                        if(cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                            checkedTCUI.Find(t => t.tc.Name == "Оптико-электронные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter++;
                        }
                        else
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Вибрационные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Оптико-электронные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                        }
                        break;
                    }
                case "cbZazem":
                    {
                        if (cb.Checked)
                        {
                            checkedTCUI.Find(t => t.tc.Name == "Воздушные" && t.tc.TCUIType.Name == "Каналы утечки акустической (речевой) информации").counter--;
                            checkedTCUI.Find(t => t.tc.Name == "Электрические" && t.tc.TCUIType.Name == "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)").counter--;
                        }
                        else
                        {

                        }
                        break;
                    }
            }
            #endregion
            updateListOfControls();
        }

        public void enterAtPageAbilsOfIntruder(object sender, EventArgs e)
        {
            mf.tabControlTCUI.TabPages["tabPageIntrAbil"].Controls.Clear();
            if (controlsIAC.Count == 0)
                mf.tabControlTCUI.TabPages["tabPageIntrAbil"].Controls.Add(mf.lbAvilitiesInfo);
            for (int i = 0; i < controlsIAC.Count; i++)
            {
                controlsIAC[i].Location = new System.Drawing.Point { X = 15, Y = (i * (controlsIAC[i].Height + 15)) + 15 };
                mf.tabControlTCUI.TabPages["tabPageIntrAbil"].Controls.Add(controlsIAC[i]);
            }
        }

        public void updateListOfControls()
        {
            foreach (checkedTCUI ctc in checkedTCUI)
            {
                if (ctc.counter > 0)
                {
                    foreach (TCUIThreat tcthreat in ctc.tc.TCUIThreats)
                    {
                        IntruderAbilityControl iac = new IntruderAbilityControl(tcthreat.Name, ctc.tc.Name, ctc.tc.TCUIType.Name, mf);
                        if (controlsIAC.Find(t => t.threatName == tcthreat.Name) == null)
                        {
                            listOfTCUIThreats.Add(tcthreat);
                            controlsIAC.Add(iac);
                        }
                    }
                }
                else
                {
                    foreach (TCUIThreat tct in ctc.tc.TCUIThreats)
                    {
                        int index = controlsIAC.FindIndex(t => t.threatName == tct.Name && t.TCUI == ctc.tc.Name);
                        if (index != -1)
                        {
                            controlsIAC.RemoveAt(index);
                            listOfTCUIThreats.RemoveAt(index);
                        }
                    }
                }
            }
        }

        public string getTypeOfTCUI(string TCUI)
        {
            if (TCUI.Contains("Voice"))
                return "Каналы утечки акустической (речевой) информации";
            if (TCUI.Contains("Pereh"))
                return "Каналы перехвата информации при ее передаче по каналам связи";
            if (TCUI.Contains("PEMIN"))
                return "Каналы побочных электромагнитных излучений и наводок (ПЭМИН)";
            return "Каналы утечки видовой информации";
        }

        public void enterTabPageThreatsList(object sender, EventArgs e)
        {
            mf.dgvActualTCUIThreats.Rows.Clear();
            
            ThreatSource ts =new ThreatSource { Potencial = getMaxIntrPot(IS.listOfSources) };

            if(ts.Potencial==-1)
            {
                MessageBox.Show("Перед определением актуальных угроз утечки информации по техническим каналам выберите нарушителя в соответствующей вкладке.");
                mf.tabControlTCUI.SelectedTab = mf.tabControlTCUI.TabPages[0];
                mf.treeView.SelectedNode = mf.returnTreeNode("tnIntruder");
                return;
            }

            List<TCUIThreat> actualTCUIThreats = new List<TCUIThreat>();
            foreach (IntruderAbilityControl iac in controlsIAC)
            {
                iac.updateIac();
                int intrPot = (int)iac.intrud;
                if ((iac.threatValue >= 10 && iac.Checked) && iac.abilityOfRealization != "" && iac.damage != "")
                {
                    if (iac.damage == "Высокая" && intrPot <= ts.Potencial)
                        actualTCUIThreats.Add(listOfTCUIThreats.Find(t => t.Name == iac.threatName));

                    if (iac.damage == "Средняя" && (iac.abilityOfRealization == "Средняя" || iac.abilityOfRealization == "Высокая") && intrPot <= ts.Potencial)
                        actualTCUIThreats.Add(listOfTCUIThreats.Find(t => t.Name == iac.threatName));

                    if (iac.damage == "Низкая" && iac.abilityOfRealization == "Высокая" && intrPot <= ts.Potencial)
                        actualTCUIThreats.Add(listOfTCUIThreats.Find(t => t.Name == iac.threatName));
                }
            }

            foreach (TCUIThreat tct in actualTCUIThreats)
            {
                mf.dgvActualTCUIThreats.Rows.Add(tct.Identificator + " " + tct.Name, tct.Description);
            }

            bool fullList = true;
            foreach (IntruderAbilityControl iac in controlsIAC)
            {
                if (!iac.Checked)
                    fullList = false;
            }

            if (fullList)
                if (mf.dgvActualTCUIThreats.Rows.Count == 0)
                    mf.lbTCUIInfo.Text = "Угрозы утечки информации по техническим каналам не актуальны для информационной системы.";
                else
                    mf.lbTCUIInfo.Text = "Угрозы утечки информации по техническим каналам в списке являются актуальными для информационной системы.";
            else
                mf.lbTCUIInfo.Text = "Для определения списка актуальных угроз утечки по техническим каналам, выберите все поля на предыдущих вкладках.";
            setDGVHeight();
        }

        public void setDGVHeight()
        {
            foreach (DataGridViewRow dgvr in mf.dgvActualTCUIThreats.Rows)
            {
                int index = dgvr.Index;
                int d = mf.dgvActualTCUIThreats.Rows[index].GetPreferredHeight(index, DataGridViewAutoSizeRowMode.AllCells, true);
                dgvr.Height = d;
            }
            mf.dgvActualTCUIThreats.Height = mf.dgvActualTCUIThreats.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + mf.dgvActualTCUIThreats.ColumnHeadersHeight;
        }

        public int getMaxIntrPot(List<ThreatSource> list)
        {
            int max = -1;
            if (list.Count != 0)
            {
                foreach (ThreatSource ts in list)
                {
                    if (ts.Potencial > max && ts.Potencial != 3)
                        max = ts.Potencial;
                }
                return max;
            }
            else
                return -1;
        }
    }
}
