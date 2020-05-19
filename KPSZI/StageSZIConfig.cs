using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;
using KPSZI.Model;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Security.Cryptography;

namespace KPSZI
{
    class StageSZIConfig : Stage
    {
        /// <summary>
        /// Параметры настроек СЗИ от НСД
        /// </summary>
        class SZIConfigInfo
        {
            /// <summary>
            /// Описание параметра настроек
            /// </summary>
            public string Name { get; private set; }
            /// <summary>
            /// значение параметра настроек
            /// </summary>
            public string Value { get; private set; }

            /// <summary>
            /// словарь описаний для параметров настроек
            /// </summary>
            public Dictionary<string, string> ParametersDescription
            {
                get
                {
                    return new Dictionary<string, string>()
                    {
                        ["account.minPwLen"] = "Минимальная допустимая длина пароля",
                        ["logon.maxErrorCount"] = "Максимальное количество неудачных попыток входа",
                        ["logon.lockTimeout"] = "Время блокировки при неправильном вводе пароля (мин)"
                    };
                }
            }

            public SZIConfigInfo(string parameter)
            {
                string[] splittedParam = parameter.Split('=');
                Name = splittedParam[0];
                Value = splittedParam[1];

                if (ParametersDescription.TryGetValue(Name, out string val))
                {
                    Name = val;
                }
            }
        }

        /// <summary>
        /// Отличия в конфигурационных файлах СЗИ от НСД
        /// </summary>
        class SZIDifference
        {
            /// <summary>
            /// значение в эталонном файле
            /// </summary>
            public SZIConfigInfo StandardParameter { get; private set; }
            /// <summary>
            /// значение в кастомном файле
            /// </summary>
            public SZIConfigInfo CustomParameter { get; private set; }
            public SZIDifference(SZIConfigInfo std, SZIConfigInfo cust)
            {
                StandardParameter = std;
                CustomParameter = cust;
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
        List<SZIConfigInfo> standardConfigList;

        /// <summary>
        /// параметры, описанные в кастомном файле
        /// </summary>
        List<SZIConfigInfo> customConfigList;

        /// <summary>
        /// список отличий в конфигурационных файлах
        /// </summary>
        List<SZIDifference> confDifferences;

        public StageSZIConfig(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
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
            mf.btnSelectStandardSZIConfig.Click += new EventHandler(btnSelectStandardDLC_Click);
            mf.btnSelectCustomSZIConfig.Click += new EventHandler(btnSelectCustomDLC_Click);
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
        List<SZIConfigInfo> GetConfig(string path)
        {
            StreamReader sr = new StreamReader(path);
            string line;
            List<SZIConfigInfo> result = new List<SZIConfigInfo>();
            
            for (int i = 0; i < 8; i++)
            {
                line = sr.ReadLine();
            }
            
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Contains('='))
                {
                    SZIConfigInfo configInfoLine = new SZIConfigInfo(line);
                    result.Add(configInfoLine);
                }
            }
            sr.Close();
            return result;
        }

        /// <summary>
        /// заполняет список параметров конфигурации
        /// </summary>
        /// <param name="list"></param>
        /// <param name="path"></param>
        void SetConfigList(ref List<SZIConfigInfo> list, string path)
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
            mf.lvSZIDifferences.Clear();
            if (confDifferences.Count > 0)
            {
                MessageBox.Show("Обнаружено несоответствие файла настроек эталонным значениям!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);

                mf.lvSZIDifferences.Columns.Add("Параметр", 346);
                mf.lvSZIDifferences.Columns.Add("Ожидаемое значение", 225);
                mf.lvSZIDifferences.Columns.Add("Значение в конфигурационном файле", 225);

                foreach (SZIDifference diff in confDifferences)
                {
                    string description = diff.StandardParameter.Name;
                    string standardValue = diff.StandardParameter.Value;
                    string customValue = diff.CustomParameter.Value;

                    ListViewItem item = new ListViewItem(description);
                    item.SubItems.Add(standardValue);
                    item.SubItems.Add(customValue);
                    mf.lvSZIDifferences.Items.Add(item);
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
            confDifferences = new List<SZIDifference>();
            int count = standardConfigList.Count;

            for (int i = 0; i < count; i++)
            {
                if (standardConfigList[i].Value != customConfigList[i].Value)
                {
                    SZIDifference diff = new SZIDifference(standardConfigList[i], customConfigList[i]);
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
