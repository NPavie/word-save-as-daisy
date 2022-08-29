using System.Windows.Forms;

namespace Daisy.SaveAsDAISY
{
    partial class Launcher
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
            this.StartConversion = new System.Windows.Forms.Button();
            this.browseFile = new System.Windows.Forms.Button();
            this.selectedFilePath = new System.Windows.Forms.TextBox();
            this.progressionText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // StartConversion
            // 
            this.StartConversion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.StartConversion.Location = new System.Drawing.Point(646, 10);
            this.StartConversion.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StartConversion.Name = "StartConversion";
            this.StartConversion.Size = new System.Drawing.Size(142, 41);
            this.StartConversion.TabIndex = 0;
            this.StartConversion.Text = "Start conversion";
            this.StartConversion.UseVisualStyleBackColor = true;
            this.StartConversion.Click += new System.EventHandler(this.launchScript_Click);
            // 
            // browseFile
            // 
            this.browseFile.Location = new System.Drawing.Point(12, 10);
            this.browseFile.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.browseFile.Name = "browseFile";
            this.browseFile.Size = new System.Drawing.Size(148, 41);
            this.browseFile.TabIndex = 1;
            this.browseFile.Text = "Choose a word document";
            this.browseFile.UseVisualStyleBackColor = true;
            this.browseFile.Click += new System.EventHandler(this.browseFile_Click);
            // 
            // selectedFilePath
            // 
            this.selectedFilePath.AllowDrop = true;
            this.selectedFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.selectedFilePath.Location = new System.Drawing.Point(166, 19);
            this.selectedFilePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.selectedFilePath.Name = "selectedFilePath";
            this.selectedFilePath.Size = new System.Drawing.Size(474, 22);
            this.selectedFilePath.TabIndex = 2;
            this.selectedFilePath.TextChanged += new System.EventHandler(this.selectedFilePath_TextChanged);
            this.selectedFilePath.DragDrop += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragDrop);
            this.selectedFilePath.DragEnter += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragEnter);
            // 
            // progressionText
            // 
            this.progressionText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressionText.Location = new System.Drawing.Point(13, 56);
            this.progressionText.Multiline = true;
            this.progressionText.Name = "progressionText";
            this.progressionText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.progressionText.Size = new System.Drawing.Size(775, 269);
            this.progressionText.TabIndex = 3;
            // 
            // Launcher
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 337);
            this.Controls.Add(this.progressionText);
            this.Controls.Add(this.selectedFilePath);
            this.Controls.Add(this.browseFile);
            this.Controls.Add(this.StartConversion);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Launcher";
            this.ShowIcon = false;
            this.Text = "Save As DAISY - Word to DTBook XML";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button StartConversion;
        private Button browseFile;
        private TextBox selectedFilePath;
        public TextBox progressionText;
    }
}