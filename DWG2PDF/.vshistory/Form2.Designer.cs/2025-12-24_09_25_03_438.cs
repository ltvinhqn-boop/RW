namespace DWG2PDF
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.FolderTreeView = new System.Windows.Forms.TreeView();
            this.FolderContentsListView = new System.Windows.Forms.ListView();
            this.ImageList = new System.Windows.Forms.ImageList(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tableLayoutPanel1.Controls.Add(this.FolderTreeView, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.FolderContentsListView, 1, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(1, 1);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(995, 798);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // FolderTreeView
            // 
            this.FolderTreeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FolderTreeView.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FolderTreeView.Location = new System.Drawing.Point(3, 3);
            this.FolderTreeView.Name = "FolderTreeView";
            this.FolderTreeView.Size = new System.Drawing.Size(292, 792);
            this.FolderTreeView.TabIndex = 0;
            this.FolderTreeView.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.FolderTreeView_BeforeExpand);
            // 
            // FolderContentsListView
            // 
            this.FolderContentsListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FolderContentsListView.HideSelection = false;
            this.FolderContentsListView.Location = new System.Drawing.Point(301, 3);
            this.FolderContentsListView.Name = "FolderContentsListView";
            this.FolderContentsListView.Size = new System.Drawing.Size(691, 792);
            this.FolderContentsListView.TabIndex = 1;
            this.FolderContentsListView.UseCompatibleStateImageBehavior = false;
            // 
            // ImageList
            // 
            this.ImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            this.ImageList.ImageSize = new System.Drawing.Size(24, 24);
            this.ImageList.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(996, 800);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Thư viện linh kiện";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TreeView FolderTreeView;
        private System.Windows.Forms.ListView FolderContentsListView;
        private System.Windows.Forms.ImageList ImageList;
    }
}