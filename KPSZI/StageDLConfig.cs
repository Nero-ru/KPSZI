using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;

namespace KPSZI
{
    class StageDLConfig : Stage
    {
        class Difference
        {
            public int Line { get; set; }
            public string StandardValue { get; set; }
            public string CustomValue { get; set; }

            /// <summary>
            /// класс, определяющий отличия в конфигурационных файлах
            /// </summary>
            /// <param name="ln">строка</param>
            /// <param name="std">значение в эталонном файле</param>
            /// <param name="cust">значение в кастомном файле</param>
            public Difference(int ln, string std, string cust)
            {
                Line = ln;
                StandardValue = std;
                CustomValue = cust;
            }
        }
        
        protected override ImageList imageListForTabPage { get; set; }

        /// <summary>
        /// путь к эталонному файлу
        /// </summary>
        string standardPath = string.Empty;

        /// <summary>
        /// путь к кастомному файлу
        /// </summary>
        string customPath = string.Empty;

        /// <summary>
        /// параметры, описанные в эталонном файле
        /// </summary>
        List<string> standardConfigList;

        /// <summary>
        /// параметры, описанные в кастомном файле
        /// </summary>
        List<string> customConfigList;

        /// <summary>
        /// список отличий в конфигурационных файлах
        /// </summary>
        List<Difference> confDifferences;

        public StageDLConfig(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {

        }

        public override void saveChanges()
        {

        }

        public override void enterTabPage()
        {
          
        }

        protected override void initTabPage()
        {
            mf.btnSelectStandardDLC.Click += new EventHandler(btnSelectStandardDLC_Click);
            mf.btnSelectCustomDLC.Click += new EventHandler(btnSelectCustomDLC_Click);
            mf.btnCompareDLConfigs.Click += new EventHandler(btnCompareDLConfigs_Click);                       
        }

        bool SetSourcePath(ref string path, OpenFileDialog ofd)
        {
            if (ofd.ShowDialog() == DialogResult.Cancel)
                return false;
            path = ofd.FileName;
            return true;
        }

        /// <summary>
        /// возвращает список параметров конфигурации
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        List<string> GetConfig(string path)
        {
            StreamReader sr = new StreamReader(path);
            string line;
            List<string> result = new List<string>();
            
            for (int i = 0; (line = sr.ReadLine()) != null; i++)
            {
                result.Add(line);
            }
            sr.Close();
            return result;
        }

        /// <summary>
        /// заполняет список параметров конфигурации
        /// </summary>
        /// <param name="list"></param>
        /// <param name="path"></param>
        void SetConfigList(ref List<string> list, string path)
        {
            list = GetConfig(path);
        }

        void btnSelectStandardDLC_Click(object sender, EventArgs e)
        {
            if (SetSourcePath(ref standardPath, new OpenFileDialog()))
            {
                mf.tbStandardDLCPath.Text = standardPath;
            }
        }

        void btnSelectCustomDLC_Click(object sender, EventArgs e)
        {
            if (SetSourcePath(ref customPath, new OpenFileDialog()))
            {
                mf.tbCustomDLCPath.Text = customPath;
            }
        }

        /// <summary>
        /// заполняет списки эталонных и кастомных параметров
        /// </summary>
        void FillConfigLists()
        {
            SetConfigList(ref standardConfigList, standardPath);
            SetConfigList(ref customConfigList, customPath);
        }

        /// <summary>
        /// заполняет ListView с отличиями
        /// </summary>
        void FillLvDLCDifferences()
        {
            if (confDifferences.Count > 0)
            {
                mf.lvDLCDifferences.Columns.Add("Строка", 56);
                mf.lvDLCDifferences.Columns.Add("Значение в эталонном файле", 370);
                mf.lvDLCDifferences.Columns.Add("Значение в файле для сравнения", 370);

                foreach (Difference d in confDifferences)
                {
                    ListViewItem item = new ListViewItem(d.Line.ToString());
                    item.SubItems.Add(d.StandardValue);
                    item.SubItems.Add(d.CustomValue);
                    mf.lvDLCDifferences.Items.Add(item);
                }
            }
            else
            {
                MessageBox.Show("Файлы идентичны", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// сравнивает списки эталонных и кастомных параметров
        /// </summary>
        void CompareConfigLists()
        {
            confDifferences = new List<Difference>();
            int count = standardConfigList.Count;

            for (int i = 8; i < count; i++)
            {
                if (standardConfigList[i] != customConfigList[i])
                {
                    Difference diff = new Difference(i + 1, standardConfigList[i], customConfigList[i]);
                    confDifferences.Add(diff);
                }
            }
            FillLvDLCDifferences();
        }

        void btnCompareDLConfigs_Click(object sender, EventArgs e)
        {
            if (standardPath != string.Empty && customPath != string.Empty)
            {
                FillConfigLists();
                CompareConfigLists();
            }
            else
            {
                MessageBox.Show("Вы не выбрали все необходимые файлы!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
