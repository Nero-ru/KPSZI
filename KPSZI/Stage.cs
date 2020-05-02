using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    // экземпляр класса Stage - конкретный этап, реализованный в отдельной вкладке интерфейса
    abstract class Stage
    {
        internal TabPage stageTab; // ссылка на вкладку 
        protected TreeNode stageNode; // ссылка на пункт в дереве
        protected string stageName; // название этапа, отображается в дереве
        protected int treeImageIcon; // иконка этапа в дереве
        protected bool isDone; // завершен ли этап
        protected bool stageAvailable;
        protected MainForm mf; // ссылка на главную форму интерфейса
        protected InformationSystem IS;
        protected abstract ImageList imageListForTabPage { get; set; }


        public Stage(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
        {
            this.stageTab = stageTab;
            this.stageNode = stageNode;
            stageName = stageNode.Text;
            stageTab.Text = stageName;
            mf = mainForm;
            this.IS = IS;

            stageNode.ImageIndex = 1;
            stageNode.SelectedImageIndex = 1;
            imageListForTabPage = new ImageList();
            imageListForTabPage.ImageSize = new System.Drawing.Size(256, 256);

            initTabPage();
        }

        // сохранение содержимого вкладки
        public abstract void saveChanges();

        public abstract void enterTabPage();

        protected abstract void initTabPage();
    }
}
