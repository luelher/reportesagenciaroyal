using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading;
using GrupoEmporium.Saint.Reportes;
using GrupoEmporium.Varias;

namespace GrupoEmporium.Saint
{
	/// <summary>
	/// Descripción breve de FormMigrar.
	/// </summary>
	public class FormMigrar : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button BtnMigrar;
		private System.Windows.Forms.Button BtnProbar;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.ProgressBar ObjProgreso;
		private System.Windows.Forms.ListBox lstdebug;
		private System.Windows.Forms.Button BtnLimpiar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private GrupoEmporium.Saint.Reportes.Saint ObjSaint;
		private Clase_SubProceso Subp;

		public FormMigrar()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.BtnMigrar = new System.Windows.Forms.Button();
			this.BtnProbar = new System.Windows.Forms.Button();
			this.panel1 = new System.Windows.Forms.Panel();
			this.BtnLimpiar = new System.Windows.Forms.Button();
			this.ObjProgreso = new System.Windows.Forms.ProgressBar();
			this.panel2 = new System.Windows.Forms.Panel();
			this.lstdebug = new System.Windows.Forms.ListBox();
			this.panel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnMigrar
			// 
			this.BtnMigrar.Location = new System.Drawing.Point(168, 32);
			this.BtnMigrar.Name = "BtnMigrar";
			this.BtnMigrar.Size = new System.Drawing.Size(128, 32);
			this.BtnMigrar.TabIndex = 0;
			this.BtnMigrar.Text = "Comenzar a Migrar";
			this.BtnMigrar.Click += new System.EventHandler(this.BtnMigrar_Click);
			// 
			// BtnProbar
			// 
			this.BtnProbar.Location = new System.Drawing.Point(40, 32);
			this.BtnProbar.Name = "BtnProbar";
			this.BtnProbar.Size = new System.Drawing.Size(80, 32);
			this.BtnProbar.TabIndex = 1;
			this.BtnProbar.Text = "Probar Conexión";
			this.BtnProbar.Click += new System.EventHandler(this.BtnProbar_Click);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.BtnLimpiar);
			this.panel1.Controls.Add(this.ObjProgreso);
			this.panel1.Controls.Add(this.BtnProbar);
			this.panel1.Controls.Add(this.BtnMigrar);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(656, 136);
			this.panel1.TabIndex = 2;
			// 
			// BtnLimpiar
			// 
			this.BtnLimpiar.Location = new System.Drawing.Point(280, 80);
			this.BtnLimpiar.Name = "BtnLimpiar";
			this.BtnLimpiar.Size = new System.Drawing.Size(56, 24);
			this.BtnLimpiar.TabIndex = 3;
			this.BtnLimpiar.Text = "Limpiar";
			this.BtnLimpiar.Click += new System.EventHandler(this.BtnLimpiar_Click);
			// 
			// ObjProgreso
			// 
			this.ObjProgreso.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.ObjProgreso.Location = new System.Drawing.Point(0, 113);
			this.ObjProgreso.Name = "ObjProgreso";
			this.ObjProgreso.Size = new System.Drawing.Size(656, 23);
			this.ObjProgreso.TabIndex = 2;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.lstdebug);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(0, 136);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(656, 222);
			this.panel2.TabIndex = 3;
			// 
			// lstdebug
			// 
			this.lstdebug.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lstdebug.Location = new System.Drawing.Point(0, 0);
			this.lstdebug.Name = "lstdebug";
			this.lstdebug.Size = new System.Drawing.Size(656, 212);
			this.lstdebug.TabIndex = 0;
			// 
			// FormMigrar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 358);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.panel1);
			this.Name = "FormMigrar";
			this.Text = "Migración Saint ---> Profit";
			this.panel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnProbar_Click(object sender, System.EventArgs e)
		{
			ProbarConexion();
		}

		private void BtnLimpiar_Click(object sender, System.EventArgs e)
		{
			lstdebug.Items.Clear();
		}

		private void ProbarConexion()
		{
			ObjSaint = new GrupoEmporium.Saint.Reportes.Saint();
			try{ObjSaint.Conexion_Profit.Open();}
			catch{}
			
			lstdebug.Items.Add("Probando Conexion con BD......");

			if(ObjSaint.Conexion.State == System.Data.ConnectionState.Open) lstdebug.Items.Add("BD Saint.................OK!");
			else lstdebug.Items.Add("BD Saint..............ERROR!");

			if(ObjSaint.Conexion_Profit.State == System.Data.ConnectionState.Open) lstdebug.Items.Add("BD Profit.................OK!");
			else lstdebug.Items.Add("BD Profit..............ERROR!");

		}

		public void Migrar()
		{
			GrupoEmporium.Profit.Reportes.Profit S = new GrupoEmporium.Profit.Reportes.Profit();
			lstdebug.Items.Clear();
			lstdebug.Items.Add("Iniciando.....");

			S.MigrarProfit(S.DocumentosSaint(lstdebug),lstdebug);

			lstdebug.Items.Add("Migracion Terminada");
			lstdebug.Refresh();
			lstdebug.SendToBack();

		}


		private void BtnMigrar_Click(object sender, System.EventArgs e)
		{
			if(ObjSaint.Conexion.State != System.Data.ConnectionState.Open || ObjSaint.Conexion_Profit.State != System.Data.ConnectionState.Open)
				ProbarConexion();

			if (ObjSaint.Conexion.State == System.Data.ConnectionState.Open && ObjSaint.Conexion_Profit.State == System.Data.ConnectionState.Open)
			{
				ThreadStart TS = new ThreadStart(Migrar);
				
				Subp = new Clase_SubProceso(ref TS,false);

				Subp.Iniciar();

				//Migrar();

			}
			else
			{
				lstdebug.Items.Add("No hay Conexión..............ERROR!");
			}

		}


	}
}
