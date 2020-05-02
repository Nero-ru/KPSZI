using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace KPSZI
{
    class StageAccessMatrix : Stage
    {
        DataGridView accessMatrixRules;
        DataGridView accessMatrixResources;
        Button addSubject;
        TextBox tbNameOfObject;
        TextBox tbNameOfSubject;
        Label labAddSubject;
        Label labAddObject;
        Button addObject;
        Label labDeleteInfo;
        TextBox tbNameOfResource;
        Label labAddResource;
        Button addResource;
        Form dialogForm;
        protected override ImageList imageListForTabPage { get; set; }

        public StageAccessMatrix(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {
            
        }

        protected override void initTabPage()
        {
            stageTab.AutoScroll = true;

            accessMatrixRules = new DataGridView { Name = "Rules" };
            addObject = new Button();
            addSubject = new Button();
            labAddSubject = new Label();
            labAddObject = new Label();
            tbNameOfSubject = new TextBox();
            tbNameOfObject = new TextBox();
            tbNameOfResource = new TextBox();
            labAddResource = new Label();
            addResource = new Button();

            labAddSubject.Location = new Point { X = 15, Y = 15 };
            labAddSubject.Text = "Добавить субъект доступа";
            labAddSubject.Width = 150;

            tbNameOfSubject.Location = new System.Drawing.Point { X = 15, Y = 30 };
            tbNameOfSubject.Size = new System.Drawing.Size { Width = 115 };
            tbNameOfSubject.KeyPress += new KeyPressEventHandler(enterPressed);
            tbNameOfSubject.Name = "Subject";
            tbNameOfSubject.TabIndex = 100;

            addSubject.Location = new System.Drawing.Point { X = 135, Y = 29 };
            addSubject.Click += new System.EventHandler(addSubject_click);
            addSubject.Text = "Добавить субъект доступа";
            addSubject.TabIndex = 101;

            labDeleteInfo = new Label { Text = "Внимание! \nДля удаления объекта доступа / ресурса дважды кликните по строке в первой колонке соответствующей таблицы. \nДля удаления субъекта доступа дважды кликните по заголовку колонки любой из таблиц. Для заполнения прав доступа субъекта к ресурсу кликните на соответствующей ячейке в нижней таблице." };
            labDeleteInfo.Location = new Point { X = 15, Y = 40 };
            labDeleteInfo.TextAlign = ContentAlignment.MiddleLeft;
            labDeleteInfo.Size = new Size { Height = 100, Width = 400 };

            #region Ресурсы

            accessMatrixResources = new DataGridView { Name = "Resources" };
            accessMatrixResources.Scroll += new ScrollEventHandler(scrollMatrix);
            accessMatrixRules.Scroll += new ScrollEventHandler(scrollMatrix);

            labAddResource.Text = "Добавить объект доступа (информационный ресурс)";
            labAddResource.Location = new Point { X = 425, Y = 540 };
            labAddResource.Size = new Size { Height = 50, Width = 150 };

            addResource.Location = new System.Drawing.Point { X = 545, Y = 569 };
            addResource.Click += new System.EventHandler(addResource_click);
            addResource.Text = "Добавить объект доступа";
            addResource.TabIndex = 106;

            tbNameOfResource.Location = new System.Drawing.Point { X = 425, Y = 570 };
            tbNameOfResource.Size = new System.Drawing.Size { Width = 115 };
            tbNameOfResource.KeyPress += new KeyPressEventHandler(enterPressed);
            tbNameOfResource.Name = "Resource";
            tbNameOfResource.TabIndex = 105;

            DataGridViewTextBoxColumn dgc2 = new DataGridViewTextBoxColumn();
            dgc2.HeaderText = "Информационный ресурс";
            dgc2.Width = 120;
            dgc2.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgc2.Frozen = true;
            dgc2.ReadOnly = true;
            accessMatrixResources.CellClick += new DataGridViewCellEventHandler(cellResource_click);
            accessMatrixResources.CellDoubleClick += new DataGridViewCellEventHandler(deleteResourceDoubleClick);

            accessMatrixResources.Columns.Add(dgc2);

            accessMatrixResources.DefaultCellStyle.SelectionBackColor = Color.White;
            accessMatrixResources.DefaultCellStyle.SelectionForeColor = Color.Black;
            accessMatrixResources.BackgroundColor = Color.White;
            accessMatrixResources.ColumnHeaderMouseDoubleClick += new DataGridViewCellMouseEventHandler(deleteSubject_doubleClick);
            accessMatrixResources.AllowUserToAddRows = false;
            accessMatrixResources.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            accessMatrixResources.ScrollBars = ScrollBars.Both;
            accessMatrixResources.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            accessMatrixResources.Width = 400;
            accessMatrixResources.Height = 400;
            accessMatrixResources.Location = new System.Drawing.Point { X = 15, Y = 540 };
            accessMatrixResources.MultiSelect = false;
            accessMatrixResources.SelectionMode = DataGridViewSelectionMode.CellSelect;
            accessMatrixResources.RowHeadersVisible = false;
            accessMatrixResources.AllowUserToResizeRows = false;
            accessMatrixResources.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            accessMatrixResources.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            accessMatrixResources.AllowUserToResizeColumns = false;
            accessMatrixResources.RowsAdded += new DataGridViewRowsAddedEventHandler(refresh_Grid);
            #endregion

            #region Права и привилегии
            labAddObject.Text = "Добавить объект доступа (право или привилегию)";
            labAddObject.Location = new Point { X = 420, Y = 130 };
            labAddObject.Size = new Size { Height = 50, Width = 130 };
            labAddObject.Width = 150;

            addObject.Location = new System.Drawing.Point { X = 540, Y = 158 };
            addObject.Click += new System.EventHandler(addObject_click);
            addObject.Text = "Добавить объект доступа";
            addObject.TabIndex = 103;

            tbNameOfObject.Location = new System.Drawing.Point { X = 420, Y = 160 };
            tbNameOfObject.Size = new System.Drawing.Size { Width = 115 };
            tbNameOfObject.KeyPress += new KeyPressEventHandler(enterPressed);
            tbNameOfObject.Name = "Object";
            tbNameOfObject.TabIndex = 102;
            
            DataGridViewTextBoxColumn dgc1 = new DataGridViewTextBoxColumn();
            dgc1.HeaderText = "Привилегии и права субъектов к объектам доступа";
            dgc1.Width = 120;
            dgc1.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgc1.Frozen = true;
            dgc1.ReadOnly = true;
            accessMatrixRules.CellDoubleClick += new DataGridViewCellEventHandler(deleteObject_click);
            accessMatrixRules.Columns.Add(dgc1);

            accessMatrixRules.DefaultCellStyle.SelectionBackColor = Color.White;
            accessMatrixRules.DefaultCellStyle.SelectionForeColor = Color.Black;
            accessMatrixRules.BackgroundColor = Color.White;
            accessMatrixRules.ColumnHeaderMouseDoubleClick += new DataGridViewCellMouseEventHandler(deleteSubject_doubleClick);
            accessMatrixRules.AllowUserToAddRows = false;
            accessMatrixRules.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            accessMatrixRules.ScrollBars = ScrollBars.Both;
            accessMatrixRules.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            accessMatrixRules.Width = 400;
            accessMatrixRules.Height = 400;
            accessMatrixRules.Location = new System.Drawing.Point { X = 15, Y = 130 };
            accessMatrixRules.MultiSelect = false;
            accessMatrixRules.SelectionMode = DataGridViewSelectionMode.CellSelect;
            accessMatrixRules.RowHeadersVisible = false;
            accessMatrixRules.AllowUserToResizeRows = false;
            accessMatrixRules.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            accessMatrixRules.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            accessMatrixRules.AllowUserToResizeColumns = false;
            accessMatrixRules.RowsAdded += new DataGridViewRowsAddedEventHandler(refresh_Grid);
            #endregion
        }

        public void InitDialogForm()
        {
            System.Windows.Forms.CheckBox checkBoxRead;
            System.Windows.Forms.CheckBox checkBoxWrite;
            System.Windows.Forms.CheckBox checkBoxAdd;
            System.Windows.Forms.CheckBox checkBoxDelete;
            System.Windows.Forms.CheckBox checkBoxExec;
            System.Windows.Forms.CheckBox checkBoxSZI;
            System.Windows.Forms.Label labelInfo;
            System.Windows.Forms.Button buttonAcceptPermissions;

            dialogForm = new Form();
            dialogForm.Size = new Size { Height = 235, Width = 550 };
            dialogForm.Icon = KPSZI.Properties.Resources.mf;
            dialogForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialogForm.MaximizeBox = false;
            dialogForm.MinimizeBox = false;

            checkBoxRead = new System.Windows.Forms.CheckBox();
            checkBoxWrite = new System.Windows.Forms.CheckBox();
            checkBoxAdd = new System.Windows.Forms.CheckBox();
            checkBoxDelete = new System.Windows.Forms.CheckBox();
            checkBoxExec = new System.Windows.Forms.CheckBox();
            checkBoxSZI = new System.Windows.Forms.CheckBox();
            labelInfo = new System.Windows.Forms.Label();
            buttonAcceptPermissions = new System.Windows.Forms.Button();
            // 
            // checkBoxRead
            // 
            checkBoxRead.Location = new System.Drawing.Point(43, 39);
            checkBoxRead.Name = "checkBoxRead";
            checkBoxRead.Size = new System.Drawing.Size(157, 56);
            checkBoxRead.TabIndex = 0;
            checkBoxRead.Text = "R - разрешение на открытие файлов только для чтения";
            checkBoxRead.UseVisualStyleBackColor = true;
            // 
            // checkBoxWrite
            // 
            checkBoxWrite.Location = new System.Drawing.Point(43, 101);
            checkBoxWrite.Name = "checkBoxWrite";
            checkBoxWrite.Size = new System.Drawing.Size(140, 56);
            checkBoxWrite.TabIndex = 1;
            checkBoxWrite.Text = "W - разрешение на открытие файлов для записи";
            checkBoxWrite.UseVisualStyleBackColor = true;
            // 
            // checkBoxAdd
            // 
            checkBoxAdd.Location = new System.Drawing.Point(206, 39);
            checkBoxAdd.Name = "checkBoxAdd";
            checkBoxAdd.Size = new System.Drawing.Size(142, 56);
            checkBoxAdd.TabIndex = 2;
            checkBoxAdd.Text = "A - разрешение на создание файлов на диске/создание таблиц в БД";
            checkBoxAdd.UseVisualStyleBackColor = true;
            // 
            // checkBoxDelete
            // 
            checkBoxDelete.Location = new System.Drawing.Point(206, 101);
            checkBoxDelete.Name = "checkBoxDelete";
            checkBoxDelete.Size = new System.Drawing.Size(142, 56);
            checkBoxDelete.TabIndex = 3;
            checkBoxDelete.Text = "D - разрешение на удаление файлов/записи в БД";
            checkBoxDelete.UseVisualStyleBackColor = true;
            // 
            // checkBoxExec
            // 
            checkBoxExec.Location = new System.Drawing.Point(354, 39);
            checkBoxExec.Name = "checkBoxExec";
            checkBoxExec.Size = new System.Drawing.Size(119, 56);
            checkBoxExec.TabIndex = 4;
            checkBoxExec.Text = "Х - разрешение на запуск программ";
            checkBoxExec.UseVisualStyleBackColor = true;
            // 
            // checkBoxSZI
            // 
            checkBoxSZI.Location = new System.Drawing.Point(354, 101);
            checkBoxSZI.Name = "checkBoxSZI";
            checkBoxSZI.Size = new System.Drawing.Size(119, 56);
            checkBoxSZI.TabIndex = 5;
            checkBoxSZI.Text = "S - разрешение на настройку средств защиты";
            checkBoxSZI.UseVisualStyleBackColor = true;
            // 
            // labelInfo
            // 
            labelInfo.AutoSize = true;
            labelInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelInfo.Location = new System.Drawing.Point(40, 9);
            labelInfo.Name = "labelInfo";
            labelInfo.Size = new System.Drawing.Size(467, 17);
            labelInfo.TabIndex = 6;
            labelInfo.Text = "Выберите права доступа к информационным ресурсам для субъекта";
            // 
            // buttonAcceptPermissions
            // 
            buttonAcceptPermissions.Location = new System.Drawing.Point(398, 160);
            buttonAcceptPermissions.Name = "buttonAcceptPermissions";
            buttonAcceptPermissions.Size = new System.Drawing.Size(75, 23);
            buttonAcceptPermissions.TabIndex = 7;
            buttonAcceptPermissions.Text = "Принять";
            buttonAcceptPermissions.UseVisualStyleBackColor = true;
            buttonAcceptPermissions.Click += new EventHandler(acceptPermissionsClick);
            // 
            // Form1
            // 

            dialogForm.Controls.Add(buttonAcceptPermissions);
            dialogForm.Controls.Add(labelInfo);
            dialogForm.Controls.Add(checkBoxSZI);
            dialogForm.Controls.Add(checkBoxExec);
            dialogForm.Controls.Add(checkBoxDelete);
            dialogForm.Controls.Add(checkBoxAdd);
            dialogForm.Controls.Add(checkBoxWrite);
            dialogForm.Controls.Add(checkBoxRead);
            dialogForm.Controls.Add(new Label { Name = "RowIndex", Visible = false });
            dialogForm.Controls.Add(new Label { Name = "ColumnIndex", Visible = false });
        }

        public override void enterTabPage()
        {
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(accessMatrixRules);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(accessMatrixResources);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(addSubject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(tbNameOfSubject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(addObject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(tbNameOfObject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(labAddObject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(labAddSubject);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(labDeleteInfo);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(addResource);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(tbNameOfResource);
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(labAddResource);
        }

        public override void saveChanges()
        {
                        
        }

        #region Обработчики событий
        //прокрутка матрицы
        public void scrollMatrix(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                if (((DataGridView)sender).Name == "Rules")
                    accessMatrixResources.HorizontalScrollingOffset = e.NewValue;
                if (((DataGridView)sender).Name == "Resources")
                    accessMatrixRules.HorizontalScrollingOffset = e.NewValue;
            }
        }


        // по нажатию ентер из текстбокса добавляется объект/субъект/ресурс
        public void enterPressed(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                switch (((TextBox)sender).Name)
                {
                    case "Subject":
                        addSubject_click(null, null);
                        break;
                    case "Object":
                        addObject_click(null, null);
                        break;
                    case "Resource":
                        addResource_click(null, null);
                        break;

                }
            }
            else
                mf.PressTheKey(e);
        }

        public void refresh_Grid(object sender, DataGridViewRowsAddedEventArgs e)
        {
            accessMatrixRules.Refresh();
            accessMatrixResources.Refresh();
        }

        //добавляем ресурс в правую таблицу
        public void addResource_click(object sender, EventArgs e)
        {
            if (tbNameOfResource.Text == "")
            {
                MessageBox.Show("Введите название ресурса!");
                tbNameOfResource.Focus();
                return;
            }
            DataGridViewRow dgr = new DataGridViewRow();
            int rowIndex = accessMatrixResources.Rows.Add();
            accessMatrixResources.Refresh();
            accessMatrixResources.Rows[rowIndex].Cells[0].Value = tbNameOfResource.Text;
            tbNameOfResource.Text = "";
            tbNameOfResource.Focus();
        }

        //добавляем субъект в обе таблицы
        public void addSubject_click(object sender, EventArgs e)
        {
            if (tbNameOfSubject.Text == "")
            {
                MessageBox.Show("Введите название субъекта доступа!");
                tbNameOfSubject.Focus();
                return;
            }

            DataGridViewColumn dgc;
            dgc = new DataGridViewCheckBoxColumn();
            dgc.SortMode = DataGridViewColumnSortMode.NotSortable;
            accessMatrixRules.Columns.Add(dgc);
            dgc.Width = 120;
            dgc.HeaderCell.Value = tbNameOfSubject.Text;

            DataGridViewColumn dgc2;
            dgc2 = new DataGridViewTextBoxColumn();
            dgc2.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgc2.ReadOnly = true;
            accessMatrixResources.Columns.Add(dgc2);
            dgc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgc2.Width = 120;
            dgc2.HeaderCell.Value = tbNameOfSubject.Text;

            accessMatrixRules.Refresh();
            accessMatrixResources.Refresh();
            tbNameOfSubject.Text = "";
            tbNameOfSubject.Focus();
        }

        //добавляем объект в левую таблицу
        public void addObject_click(object sender, EventArgs e)
        {
            if (tbNameOfObject.Text == "")
            {
                MessageBox.Show("Введите название объекта доступа!");
                tbNameOfObject.Focus();
                return;
            }
            DataGridViewRow dgr = new DataGridViewRow();
            int rowIndex = accessMatrixRules.Rows.Add();
            accessMatrixRules.Refresh();
            accessMatrixRules.Rows[rowIndex].Cells[0].Value = tbNameOfObject.Text;
            tbNameOfObject.Text = "";
            tbNameOfObject.Focus();
        }

        //Удаляем объект дабл-кликом в таблице объектов
        public void deleteObject_click(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                accessMatrixRules.Rows.RemoveAt(e.RowIndex);
            }
            accessMatrixRules.Refresh();
        }

        //клик по кнопке в диалоговой форме - записываем в таблицу ресурсов разрешения
        public void acceptPermissionsClick(object sender, EventArgs e)
        {
            string Rules = "";
            int RowIndex;
            int.TryParse(dialogForm.Controls["RowIndex"].Text, out RowIndex);
            int ColumnIndex;
            int.TryParse(dialogForm.Controls["ColumnIndex"].Text, out ColumnIndex);

            if (((CheckBox)(dialogForm.Controls["checkBoxRead"])).Checked)
                Rules += "R";
            if (((CheckBox)(dialogForm.Controls["checkBoxWrite"])).Checked)
                Rules += "W";
            if (((CheckBox)(dialogForm.Controls["checkBoxAdd"])).Checked)
                Rules += "A";
            if (((CheckBox)(dialogForm.Controls["checkBoxExec"])).Checked)
                Rules += "X";
            if (((CheckBox)(dialogForm.Controls["checkBoxDelete"])).Checked)
                Rules += "D";
            if (((CheckBox)(dialogForm.Controls["checkBoxSZI"])).Checked)
                Rules += "S";

            accessMatrixResources.Rows[RowIndex].Cells[ColumnIndex].Value = Rules;

            dialogForm.Close();
            dialogForm.Dispose();
        }

        //Открываем форму и заполняем ее, если в таблице было что-то отмечено
        public void openPermissionsDialog(int RowIndex, int ColumnIndex)
        {
            InitDialogForm();

            string Rules = (accessMatrixResources.Rows[RowIndex].Cells[ColumnIndex].Value != null) ? accessMatrixResources.Rows[RowIndex].Cells[ColumnIndex].Value.ToString() : "";
            if (Rules.Contains("R"))
                ((CheckBox)(dialogForm.Controls["checkBoxRead"])).Checked = true;
            if (Rules.Contains("W"))
                ((CheckBox)(dialogForm.Controls["checkBoxWrite"])).Checked = true;
            if (Rules.Contains("A"))
                ((CheckBox)(dialogForm.Controls["checkBoxAdd"])).Checked = true;
            if (Rules.Contains("X"))
                ((CheckBox)(dialogForm.Controls["checkBoxExec"])).Checked = true;
            if (Rules.Contains("D"))
                ((CheckBox)(dialogForm.Controls["checkBoxDelete"])).Checked = true;
            if (Rules.Contains("S"))
                ((CheckBox)(dialogForm.Controls["checkBoxSZI"])).Checked = true;

            string Subject = accessMatrixResources.Columns[ColumnIndex].HeaderText;
            string Resource = accessMatrixResources.Rows[RowIndex].Cells[0].Value.ToString();
            dialogForm.Text = "Субъект доступа - " + Subject + ", ресурс - " + Resource;
            dialogForm.Controls["RowIndex"].Text = RowIndex.ToString();
            dialogForm.Controls["ColumnIndex"].Text = ColumnIndex.ToString();
            dialogForm.ShowDialog();
        }

        //дабл-кликом по 0 ячейке в строке удаляем ресурс из второй таблицы
        public void deleteResourceDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                accessMatrixResources.Rows.RemoveAt(e.RowIndex);
            }
        }

        // Открываем форму с разрешениями доступа по клику любой ячейки второй таблицы
        public void cellResource_click(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != 0)
            {
                openPermissionsDialog(e.RowIndex, e.ColumnIndex);
            }
            accessMatrixResources.Refresh();
        }

        //Удаляем субъект дабл-кликом в обеих таблицах
        public void deleteSubject_doubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                MessageBox.Show("Нельзя удалить главную колонку!");
                return;
            }
            accessMatrixRules.Columns.RemoveAt(e.ColumnIndex);
            accessMatrixResources.Columns.RemoveAt(e.ColumnIndex);

        }
        #endregion

    }
}
