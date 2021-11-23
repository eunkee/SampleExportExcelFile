
namespace SampleExportExcelFile
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.GridViewMaterial = new System.Windows.Forms.DataGridView();
            this.m_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.m_material_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.m_receiving_count = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonExport = new System.Windows.Forms.Button();
            this.buttonLoad = new System.Windows.Forms.Button();
            this.ProgressBarExport = new SampleExportExcelFile.CustomProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewMaterial)).BeginInit();
            this.SuspendLayout();
            // 
            // GridViewMaterial
            // 
            this.GridViewMaterial.AllowUserToAddRows = false;
            this.GridViewMaterial.AllowUserToDeleteRows = false;
            this.GridViewMaterial.AllowUserToResizeColumns = false;
            this.GridViewMaterial.AllowUserToResizeRows = false;
            this.GridViewMaterial.BackgroundColor = System.Drawing.Color.White;
            this.GridViewMaterial.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.GridViewMaterial.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(220)))), ((int)(((byte)(220)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Transparent;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black;
            this.GridViewMaterial.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.GridViewMaterial.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GridViewMaterial.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.m_date,
            this.m_material_name,
            this.m_receiving_count});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Transparent;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.GridViewMaterial.DefaultCellStyle = dataGridViewCellStyle2;
            this.GridViewMaterial.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.GridViewMaterial.EnableHeadersVisualStyles = false;
            this.GridViewMaterial.Location = new System.Drawing.Point(0, 63);
            this.GridViewMaterial.Margin = new System.Windows.Forms.Padding(2);
            this.GridViewMaterial.Name = "GridViewMaterial";
            this.GridViewMaterial.RowHeadersVisible = false;
            this.GridViewMaterial.RowTemplate.Height = 25;
            this.GridViewMaterial.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.GridViewMaterial.Size = new System.Drawing.Size(618, 215);
            this.GridViewMaterial.TabIndex = 1;
            this.GridViewMaterial.Click += new System.EventHandler(this.GridViewMaterial_Click);
            // 
            // m_date
            // 
            this.m_date.HeaderText = "날짜";
            this.m_date.Name = "m_date";
            this.m_date.ReadOnly = true;
            this.m_date.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.m_date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.m_date.Width = 190;
            // 
            // m_material_name
            // 
            this.m_material_name.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.m_material_name.FillWeight = 200F;
            this.m_material_name.HeaderText = "거래처";
            this.m_material_name.Name = "m_material_name";
            this.m_material_name.ReadOnly = true;
            this.m_material_name.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.m_material_name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // m_receiving_count
            // 
            this.m_receiving_count.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.m_receiving_count.HeaderText = "입고 수량";
            this.m_receiving_count.Name = "m_receiving_count";
            this.m_receiving_count.ReadOnly = true;
            this.m_receiving_count.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.m_receiving_count.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // buttonExport
            // 
            this.buttonExport.Location = new System.Drawing.Point(513, 21);
            this.buttonExport.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new System.Drawing.Size(75, 22);
            this.buttonExport.TabIndex = 0;
            this.buttonExport.Text = "Export";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new System.EventHandler(this.ButtonExport_Click);
            // 
            // buttonLoad
            // 
            this.buttonLoad.Location = new System.Drawing.Point(28, 21);
            this.buttonLoad.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(75, 22);
            this.buttonLoad.TabIndex = 2;
            this.buttonLoad.Text = "Load";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.ButtonLoad_Click);
            // 
            // ProgressBarExport
            // 
            this.ProgressBarExport.Location = new System.Drawing.Point(124, 21);
            this.ProgressBarExport.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ProgressBarExport.Name = "ProgressBarExport";
            this.ProgressBarExport.Size = new System.Drawing.Size(369, 22);
            this.ProgressBarExport.TabIndex = 0;
            this.ProgressBarExport.Text = null;
            this.ProgressBarExport.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.ProgressBarExport.TextColor = System.Drawing.SystemColors.ControlText;
            this.ProgressBarExport.TextFont = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold);
            this.ProgressBarExport.TextMargin = new System.Drawing.Point(1, 1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(618, 278);
            this.Controls.Add(this.ProgressBarExport);
            this.Controls.Add(this.buttonLoad);
            this.Controls.Add(this.GridViewMaterial);
            this.Controls.Add(this.buttonExport);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Sample Export Excel File";
            ((System.ComponentModel.ISupportInitialize)(this.GridViewMaterial)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.DataGridView GridViewMaterial;
        private System.Windows.Forms.Button buttonLoad;
        private System.Windows.Forms.DataGridViewTextBoxColumn m_date;
        private System.Windows.Forms.DataGridViewTextBoxColumn m_material_name;
        private System.Windows.Forms.DataGridViewTextBoxColumn m_receiving_count;
        private CustomProgressBar ProgressBarExport;
    }
}

