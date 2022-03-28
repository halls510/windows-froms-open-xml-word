
namespace WindowsFormsAppOpenXML
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.Txt_Name = new System.Windows.Forms.TextBox();
            this.Btn_Aplicar = new System.Windows.Forms.Button();
            this.Lbl_Name = new System.Windows.Forms.Label();
            this.Btn_Visualizar = new System.Windows.Forms.Button();
            this.Rtx_Conteudo = new System.Windows.Forms.RichTextBox();
            this.Btn_FileNameOrigin = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Txt_NameFileOrigin = new System.Windows.Forms.TextBox();
            this.Btn_FileNameDestination = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.Txt_FilePathDestination = new System.Windows.Forms.TextBox();
            this.Txt_NameFileDestination = new System.Windows.Forms.TextBox();
            this.Lbl_NameFileDestination = new System.Windows.Forms.Label();
            this.Txt_Region = new System.Windows.Forms.TextBox();
            this.Lbl_Regiao = new System.Windows.Forms.Label();
            this.Combo_Regiao = new System.Windows.Forms.ComboBox();
            this.Btn_AddRegiao = new System.Windows.Forms.Button();
            this.Lbl_AddRegiao = new System.Windows.Forms.Label();
            this.Btn_AddPessoaInRegion = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Txt_Name
            // 
            this.Txt_Name.Location = new System.Drawing.Point(203, 211);
            this.Txt_Name.Name = "Txt_Name";
            this.Txt_Name.Size = new System.Drawing.Size(239, 20);
            this.Txt_Name.TabIndex = 0;
            // 
            // Btn_Aplicar
            // 
            this.Btn_Aplicar.Location = new System.Drawing.Point(16, 532);
            this.Btn_Aplicar.Name = "Btn_Aplicar";
            this.Btn_Aplicar.Size = new System.Drawing.Size(280, 37);
            this.Btn_Aplicar.TabIndex = 3;
            this.Btn_Aplicar.Text = "Aplicar";
            this.Btn_Aplicar.UseVisualStyleBackColor = true;
            this.Btn_Aplicar.Click += new System.EventHandler(this.Btn_Aplicar_Click);
            // 
            // Lbl_Name
            // 
            this.Lbl_Name.AutoSize = true;
            this.Lbl_Name.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lbl_Name.Location = new System.Drawing.Point(12, 207);
            this.Lbl_Name.Name = "Lbl_Name";
            this.Lbl_Name.Size = new System.Drawing.Size(66, 24);
            this.Lbl_Name.TabIndex = 5;
            this.Lbl_Name.Text = "Name:";
            // 
            // Btn_Visualizar
            // 
            this.Btn_Visualizar.Location = new System.Drawing.Point(12, 590);
            this.Btn_Visualizar.Name = "Btn_Visualizar";
            this.Btn_Visualizar.Size = new System.Drawing.Size(280, 37);
            this.Btn_Visualizar.TabIndex = 9;
            this.Btn_Visualizar.Text = "Visuarlizar Arquivo";
            this.Btn_Visualizar.UseVisualStyleBackColor = true;
            this.Btn_Visualizar.Click += new System.EventHandler(this.Btn_Visualizar_Click);
            // 
            // Rtx_Conteudo
            // 
            this.Rtx_Conteudo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Rtx_Conteudo.Location = new System.Drawing.Point(573, 27);
            this.Rtx_Conteudo.Name = "Rtx_Conteudo";
            this.Rtx_Conteudo.Size = new System.Drawing.Size(603, 600);
            this.Rtx_Conteudo.TabIndex = 10;
            this.Rtx_Conteudo.Text = "";
            // 
            // Btn_FileNameOrigin
            // 
            this.Btn_FileNameOrigin.Location = new System.Drawing.Point(12, 342);
            this.Btn_FileNameOrigin.Name = "Btn_FileNameOrigin";
            this.Btn_FileNameOrigin.Size = new System.Drawing.Size(280, 40);
            this.Btn_FileNameOrigin.TabIndex = 11;
            this.Btn_FileNameOrigin.Text = "Selecionar Arquivo";
            this.Btn_FileNameOrigin.UseVisualStyleBackColor = true;
            this.Btn_FileNameOrigin.Click += new System.EventHandler(this.Btn_FileNameOrigin_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Txt_NameFileOrigin
            // 
            this.Txt_NameFileOrigin.Location = new System.Drawing.Point(311, 342);
            this.Txt_NameFileOrigin.Multiline = true;
            this.Txt_NameFileOrigin.Name = "Txt_NameFileOrigin";
            this.Txt_NameFileOrigin.Size = new System.Drawing.Size(237, 40);
            this.Txt_NameFileOrigin.TabIndex = 12;
            // 
            // Btn_FileNameDestination
            // 
            this.Btn_FileNameDestination.Location = new System.Drawing.Point(16, 471);
            this.Btn_FileNameDestination.Name = "Btn_FileNameDestination";
            this.Btn_FileNameDestination.Size = new System.Drawing.Size(280, 40);
            this.Btn_FileNameDestination.TabIndex = 13;
            this.Btn_FileNameDestination.Text = "Selecionar Destino";
            this.Btn_FileNameDestination.UseVisualStyleBackColor = true;
            this.Btn_FileNameDestination.Click += new System.EventHandler(this.Btn_FileNameDestination_Click);
            // 
            // Txt_FilePathDestination
            // 
            this.Txt_FilePathDestination.Location = new System.Drawing.Point(311, 471);
            this.Txt_FilePathDestination.Multiline = true;
            this.Txt_FilePathDestination.Name = "Txt_FilePathDestination";
            this.Txt_FilePathDestination.Size = new System.Drawing.Size(237, 40);
            this.Txt_FilePathDestination.TabIndex = 14;
            // 
            // Txt_NameFileDestination
            // 
            this.Txt_NameFileDestination.Location = new System.Drawing.Point(311, 413);
            this.Txt_NameFileDestination.Multiline = true;
            this.Txt_NameFileDestination.Name = "Txt_NameFileDestination";
            this.Txt_NameFileDestination.Size = new System.Drawing.Size(237, 40);
            this.Txt_NameFileDestination.TabIndex = 15;
            // 
            // Lbl_NameFileDestination
            // 
            this.Lbl_NameFileDestination.AutoSize = true;
            this.Lbl_NameFileDestination.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lbl_NameFileDestination.Location = new System.Drawing.Point(12, 429);
            this.Lbl_NameFileDestination.Name = "Lbl_NameFileDestination";
            this.Lbl_NameFileDestination.Size = new System.Drawing.Size(203, 24);
            this.Lbl_NameFileDestination.TabIndex = 16;
            this.Lbl_NameFileDestination.Text = "Novo nome do arquivo";
            // 
            // Txt_Region
            // 
            this.Txt_Region.Location = new System.Drawing.Point(156, 89);
            this.Txt_Region.Name = "Txt_Region";
            this.Txt_Region.Size = new System.Drawing.Size(239, 20);
            this.Txt_Region.TabIndex = 17;
            // 
            // Lbl_Regiao
            // 
            this.Lbl_Regiao.AutoSize = true;
            this.Lbl_Regiao.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lbl_Regiao.Location = new System.Drawing.Point(12, 168);
            this.Lbl_Regiao.Name = "Lbl_Regiao";
            this.Lbl_Regiao.Size = new System.Drawing.Size(180, 24);
            this.Lbl_Regiao.TabIndex = 19;
            this.Lbl_Regiao.Text = "Selecione a Região:";
            // 
            // Combo_Regiao
            // 
            this.Combo_Regiao.FormattingEnabled = true;
            this.Combo_Regiao.Location = new System.Drawing.Point(203, 168);
            this.Combo_Regiao.Name = "Combo_Regiao";
            this.Combo_Regiao.Size = new System.Drawing.Size(239, 21);
            this.Combo_Regiao.TabIndex = 20;
            // 
            // Btn_AddRegiao
            // 
            this.Btn_AddRegiao.Location = new System.Drawing.Point(422, 89);
            this.Btn_AddRegiao.Name = "Btn_AddRegiao";
            this.Btn_AddRegiao.Size = new System.Drawing.Size(100, 24);
            this.Btn_AddRegiao.TabIndex = 21;
            this.Btn_AddRegiao.Text = "Add";
            this.Btn_AddRegiao.UseVisualStyleBackColor = true;
            this.Btn_AddRegiao.Click += new System.EventHandler(this.Btn_AddRegiao_Click);
            // 
            // Lbl_AddRegiao
            // 
            this.Lbl_AddRegiao.AutoSize = true;
            this.Lbl_AddRegiao.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lbl_AddRegiao.Location = new System.Drawing.Point(12, 86);
            this.Lbl_AddRegiao.Name = "Lbl_AddRegiao";
            this.Lbl_AddRegiao.Size = new System.Drawing.Size(115, 24);
            this.Lbl_AddRegiao.TabIndex = 18;
            this.Lbl_AddRegiao.Text = "Add Região:";
            // 
            // Btn_AddPessoaInRegion
            // 
            this.Btn_AddPessoaInRegion.Location = new System.Drawing.Point(467, 189);
            this.Btn_AddPessoaInRegion.Name = "Btn_AddPessoaInRegion";
            this.Btn_AddPessoaInRegion.Size = new System.Drawing.Size(100, 24);
            this.Btn_AddPessoaInRegion.TabIndex = 22;
            this.Btn_AddPessoaInRegion.Text = "Add";
            this.Btn_AddPessoaInRegion.UseVisualStyleBackColor = true;
            this.Btn_AddPessoaInRegion.Click += new System.EventHandler(this.Btn_AddPessoaInRegion_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1188, 660);
            this.Controls.Add(this.Btn_AddPessoaInRegion);
            this.Controls.Add(this.Btn_AddRegiao);
            this.Controls.Add(this.Combo_Regiao);
            this.Controls.Add(this.Lbl_Regiao);
            this.Controls.Add(this.Lbl_AddRegiao);
            this.Controls.Add(this.Txt_Region);
            this.Controls.Add(this.Lbl_NameFileDestination);
            this.Controls.Add(this.Txt_NameFileDestination);
            this.Controls.Add(this.Txt_FilePathDestination);
            this.Controls.Add(this.Btn_FileNameDestination);
            this.Controls.Add(this.Txt_NameFileOrigin);
            this.Controls.Add(this.Btn_FileNameOrigin);
            this.Controls.Add(this.Rtx_Conteudo);
            this.Controls.Add(this.Btn_Visualizar);
            this.Controls.Add(this.Lbl_Name);
            this.Controls.Add(this.Btn_Aplicar);
            this.Controls.Add(this.Txt_Name);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Txt_Name;
        private System.Windows.Forms.Button Btn_Aplicar;
        private System.Windows.Forms.Label Lbl_Name;
        private System.Windows.Forms.Button Btn_Visualizar;
        private System.Windows.Forms.RichTextBox Rtx_Conteudo;
        private System.Windows.Forms.Button Btn_FileNameOrigin;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox Txt_NameFileOrigin;
        private System.Windows.Forms.Button Btn_FileNameDestination;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox Txt_FilePathDestination;
        private System.Windows.Forms.TextBox Txt_NameFileDestination;
        private System.Windows.Forms.Label Lbl_NameFileDestination;
        private System.Windows.Forms.TextBox Txt_Region;
        private System.Windows.Forms.Label Lbl_Regiao;
        private System.Windows.Forms.ComboBox Combo_Regiao;
        private System.Windows.Forms.Button Btn_AddRegiao;
        private System.Windows.Forms.Label Lbl_AddRegiao;
        private System.Windows.Forms.Button Btn_AddPessoaInRegion;
    }
}

