using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using GrupoEmporium.Formularios.Presentacion;
using GrupoEmporium.Mensajes;
using GrupoEmporium.Reportes;
using GrupoEmporium.Saint.Reportes;
using GrupoEmporium.Varias;

namespace GrupoEmporium.Saint
{
	/// <summary>
	/// Formulario para seleccionar el reporte de Saint que se va a generar.
	/// </summary>
	public class FormSaint : System.Windows.Forms.Form
	{

		DataTable dt;
		DataSet ds;

		private System.Windows.Forms.Button BotonCrearReporte;
		private System.Windows.Forms.DateTimePicker Fecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.LinkLabel LabelAcerca;
		private System.Windows.Forms.ComboBox ComboReporte;
		private System.Windows.Forms.DataGrid GridResultado;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Button BotonExportar;
		private System.Windows.Forms.Button BtnMigrar;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textMesDesde;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textMesHasta;
		private System.Windows.Forms.Label labelGenerando;
		private System.Windows.Forms.RadioButton rbconexion1;
		private System.Windows.Forms.RadioButton rbconexion2;
        private Button BotonExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FormSaint()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//

			ClaseDocumentosXML MiConfig = new ClaseDocumentosXML(@"configsaint.xml");

			if(MiConfig.Cargado)
			{
				if (MiConfig["Migrar"]=="False")
				{
					BtnMigrar.Visible=false;
				}
			}

		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSaint));
            this.BotonCrearReporte = new System.Windows.Forms.Button();
            this.ComboReporte = new System.Windows.Forms.ComboBox();
            this.Fecha = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.LabelAcerca = new System.Windows.Forms.LinkLabel();
            this.GridResultado = new System.Windows.Forms.DataGrid();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rbconexion2 = new System.Windows.Forms.RadioButton();
            this.rbconexion1 = new System.Windows.Forms.RadioButton();
            this.labelGenerando = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textMesHasta = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textMesDesde = new System.Windows.Forms.TextBox();
            this.BtnMigrar = new System.Windows.Forms.Button();
            this.BotonExportar = new System.Windows.Forms.Button();
            this.BotonExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.GridResultado)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BotonCrearReporte
            // 
            this.BotonCrearReporte.Location = new System.Drawing.Point(136, 128);
            this.BotonCrearReporte.Name = "BotonCrearReporte";
            this.BotonCrearReporte.Size = new System.Drawing.Size(112, 24);
            this.BotonCrearReporte.TabIndex = 0;
            this.BotonCrearReporte.Text = "Generar Reporte";
            this.BotonCrearReporte.Click += new System.EventHandler(this.BotonCrearReporte_Click);
            // 
            // ComboReporte
            // 
            this.ComboReporte.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboReporte.Items.AddRange(new object[] {
            "Cuentas Por Cobrar (Saint)",
            "Experiencia (Saint)",
            "Experiencia (Profit)",
            "Cartas Morosos 2 y 3 meses (Profit)",
            "Cartas Morosos 6 o mas  meses (Profit)",
            "Cartas Morosos Por rango de meses (Profit)",
            "Cartas Morosos Por rango de meses con Encuesta(Profit)",
            "Cartas Ultimo Aviso (Profit)",
            "Carta Legal (Profit)",
            "Listado Clientes Morosos"});
            this.ComboReporte.Location = new System.Drawing.Point(104, 40);
            this.ComboReporte.Name = "ComboReporte";
            this.ComboReporte.Size = new System.Drawing.Size(256, 21);
            this.ComboReporte.TabIndex = 1;
            // 
            // Fecha
            // 
            this.Fecha.Location = new System.Drawing.Point(104, 80);
            this.Fecha.Name = "Fecha";
            this.Fecha.Size = new System.Drawing.Size(216, 20);
            this.Fecha.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(32, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Reporte";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(32, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "Fecha";
            // 
            // LabelAcerca
            // 
            this.LabelAcerca.Location = new System.Drawing.Point(280, 0);
            this.LabelAcerca.Name = "LabelAcerca";
            this.LabelAcerca.Size = new System.Drawing.Size(64, 16);
            this.LabelAcerca.TabIndex = 5;
            this.LabelAcerca.TabStop = true;
            this.LabelAcerca.Text = "Acerca de...";
            this.LabelAcerca.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LabelAcerca_LinkClicked);
            // 
            // GridResultado
            // 
            this.GridResultado.DataMember = "";
            this.GridResultado.Dock = System.Windows.Forms.DockStyle.Fill;
            this.GridResultado.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.GridResultado.Location = new System.Drawing.Point(0, 216);
            this.GridResultado.Name = "GridResultado";
            this.GridResultado.Size = new System.Drawing.Size(394, 168);
            this.GridResultado.TabIndex = 6;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BotonExcel);
            this.panel1.Controls.Add(this.rbconexion2);
            this.panel1.Controls.Add(this.rbconexion1);
            this.panel1.Controls.Add(this.labelGenerando);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textMesHasta);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textMesDesde);
            this.panel1.Controls.Add(this.BtnMigrar);
            this.panel1.Controls.Add(this.BotonExportar);
            this.panel1.Controls.Add(this.LabelAcerca);
            this.panel1.Controls.Add(this.BotonCrearReporte);
            this.panel1.Controls.Add(this.ComboReporte);
            this.panel1.Controls.Add(this.Fecha);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(394, 216);
            this.panel1.TabIndex = 7;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // rbconexion2
            // 
            this.rbconexion2.Location = new System.Drawing.Point(280, 152);
            this.rbconexion2.Name = "rbconexion2";
            this.rbconexion2.Size = new System.Drawing.Size(104, 16);
            this.rbconexion2.TabIndex = 14;
            this.rbconexion2.Text = "Profit - Saint";
            // 
            // rbconexion1
            // 
            this.rbconexion1.Checked = true;
            this.rbconexion1.Location = new System.Drawing.Point(280, 128);
            this.rbconexion1.Name = "rbconexion1";
            this.rbconexion1.Size = new System.Drawing.Size(104, 16);
            this.rbconexion1.TabIndex = 13;
            this.rbconexion1.TabStop = true;
            this.rbconexion1.Text = "Profit Principal";
            // 
            // labelGenerando
            // 
            this.labelGenerando.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelGenerando.ForeColor = System.Drawing.Color.Red;
            this.labelGenerando.Location = new System.Drawing.Point(136, 163);
            this.labelGenerando.Name = "labelGenerando";
            this.labelGenerando.Size = new System.Drawing.Size(112, 16);
            this.labelGenerando.TabIndex = 12;
            this.labelGenerando.Text = "Generando...";
            this.labelGenerando.Visible = false;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(192, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(64, 16);
            this.label4.TabIndex = 11;
            this.label4.Text = "Mes Hasta";
            // 
            // textMesHasta
            // 
            this.textMesHasta.Location = new System.Drawing.Point(264, 104);
            this.textMesHasta.Name = "textMesHasta";
            this.textMesHasta.Size = new System.Drawing.Size(48, 20);
            this.textMesHasta.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(32, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 16);
            this.label3.TabIndex = 9;
            this.label3.Text = "Mes Desde";
            // 
            // textMesDesde
            // 
            this.textMesDesde.Location = new System.Drawing.Point(104, 104);
            this.textMesDesde.Name = "textMesDesde";
            this.textMesDesde.Size = new System.Drawing.Size(48, 20);
            this.textMesDesde.TabIndex = 8;
            // 
            // BtnMigrar
            // 
            this.BtnMigrar.Location = new System.Drawing.Point(16, 144);
            this.BtnMigrar.Name = "BtnMigrar";
            this.BtnMigrar.Size = new System.Drawing.Size(80, 32);
            this.BtnMigrar.TabIndex = 7;
            this.BtnMigrar.Text = "Migrar a Profit";
            this.BtnMigrar.Click += new System.EventHandler(this.BtnMigrar_Click);
            // 
            // BotonExportar
            // 
            this.BotonExportar.Enabled = false;
            this.BotonExportar.Location = new System.Drawing.Point(90, 182);
            this.BotonExportar.Name = "BotonExportar";
            this.BotonExportar.Size = new System.Drawing.Size(112, 24);
            this.BotonExportar.TabIndex = 6;
            this.BotonExportar.Text = "Exportar Reporte";
            this.BotonExportar.Click += new System.EventHandler(this.BotonExportar_Click);
            // 
            // BotonExcel
            // 
            this.BotonExcel.Enabled = false;
            this.BotonExcel.Location = new System.Drawing.Point(208, 182);
            this.BotonExcel.Name = "BotonExcel";
            this.BotonExcel.Size = new System.Drawing.Size(112, 24);
            this.BotonExcel.TabIndex = 15;
            this.BotonExcel.Text = "Generar Excel";
            this.BotonExcel.Click += new System.EventHandler(this.BotonExcel_Click);
            // 
            // FormSaint
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(394, 384);
            this.Controls.Add(this.GridResultado);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormSaint";
            this.Text = "Reportes Saint";
            ((System.ComponentModel.ISupportInitialize)(this.GridResultado)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Punto de entrada principal de la aplicación.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new FormSaint());
		}

		private void LabelAcerca_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Bitmap Img = new Bitmap(200,200,System.Drawing.Imaging.PixelFormat.Format24bppRgb);

            FormAcerca Acerca = new FormAcerca(ref Img,Color.YellowGreen ,Color.Black,"Aplicación para generar reportes del sistema administrativo Saint Enterprise",this.Text,"v." + Application.ProductVersion,"Grupo Emporium C.A.");
			Acerca.ShowDialog(this);
			Acerca.Dispose();
		}

		private void BotonCrearReporte_Click(object sender, System.EventArgs e)
		{
			Saint.Reportes.Saint ObjSaint;
			Profit.Reportes.Profit ObjProfit;
			labelGenerando.Visible = true;
			switch(ComboReporte.SelectedIndex)
			{
				case 0: // Cuentas por cobrar Saint
					ObjSaint = new Saint.Reportes.Saint(Fecha.Value);

					ds = ObjSaint.Reporte_Resumen_CXC();

					GridResultado.SetDataBinding(ds,"");

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        BotonExportar.Enabled = true;
                        BotonExcel.Enabled = true;
                    }
                    else {
                        BotonExportar.Enabled = false;
                        BotonExcel.Enabled = false;
                    } 

					break;
				case 1: // Experiencia Saint
					ObjSaint = new Saint.Reportes.Saint(Fecha.Value);

					dt = ObjSaint.Reporte_Experiencia();

					GridResultado.SetDataBinding(dt,"");

                    if (dt.Rows.Count > 0)
                    {
                        BotonExportar.Enabled = true;
                        BotonExcel.Enabled = true;
                    }
                    else
                    {
                        BotonExportar.Enabled = false;
                        BotonExcel.Enabled = false;
                    } 


					break;
				case 2:  // Experiencia Profit
					ObjProfit = new Profit.Reportes.Profit(Fecha.Value);

					dt = ObjProfit.Reporte_Experiencia();

					GridResultado.SetDataBinding(dt,"");

                    if (dt.Rows.Count > 0)
                    {
                        BotonExportar.Enabled = true;
                        BotonExcel.Enabled = true;
                    }
                    else
                    {
                        BotonExportar.Enabled = false;
                        BotonExcel.Enabled = false;
                    } 


					break;
				case 3: // Reporte Morosos 2 y 3 meses
					ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

					dt = ObjProfit.Reporte_Morosos(2,3);

					GridResultado.SetDataBinding(dt,"");

                    if (dt.Rows.Count > 0)
                    {
                        BotonExportar.Enabled = true;
                        BotonExcel.Enabled = true;
                    }
                    else
                    {
                        BotonExportar.Enabled = false;
                        BotonExcel.Enabled = false;
                    } 


					break;
				case 4: // Reporte Morosos 6 meses
					ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

					dt = ObjProfit.Reporte_Morosos(6,6);

					GridResultado.SetDataBinding(dt,"");

                    if (dt.Rows.Count > 0)
                    {
                        BotonExportar.Enabled = true;
                        BotonExcel.Enabled = true;
                    }
                    else
                    {
                        BotonExportar.Enabled = false;
                        BotonExcel.Enabled = false;
                    } 


					break;
				case 5: // Reporte Morosos Por Meses
					if(textMesHasta.Text!="" && textMesDesde.Text!="")
					{
						ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

						dt = ObjProfit.Reporte_Morosos(Convert.ToInt32(textMesDesde.Text),Convert.ToInt32(textMesHasta.Text));

						GridResultado.SetDataBinding(dt,"");

                        if (dt.Rows.Count > 0)
                        {
                            BotonExportar.Enabled = true;
                            BotonExcel.Enabled = true;
                        }
                        else
                        {
                            BotonExportar.Enabled = false;
                            BotonExcel.Enabled = false;
                        } 

					}else Mensaje.Error("Debe colocar el nro de los meses desde y hasta para realizar la consulta","Reportes Profit");
					break;
				case 6: // Reporte Morosos Por Meses con encuesta
					if(textMesHasta.Text!="" && textMesDesde.Text!="")
					{
						ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

						dt = ObjProfit.Reporte_Morosos(Convert.ToInt32(textMesDesde.Text),Convert.ToInt32(textMesHasta.Text));

						GridResultado.SetDataBinding(dt,"");

                        if (dt.Rows.Count > 0)
                        {
                            BotonExportar.Enabled = true;
                            BotonExcel.Enabled = true;
                        }
                        else
                        {
                            BotonExportar.Enabled = false;
                            BotonExcel.Enabled = false;
                        } 

					}
					else Mensaje.Error("Debe colocar el nro de los meses desde y hasta para realizar la consulta","Reportes Profit");
					break;
				case 7: // Reporte Morosos Por Meses + Ultimo Aviso
					if(textMesHasta.Text!="" && textMesDesde.Text!="")
					{
						ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

						dt = ObjProfit.Reporte_Morosos(Convert.ToInt32(textMesDesde.Text),Convert.ToInt32(textMesHasta.Text));

						GridResultado.SetDataBinding(dt,"");

                        if (dt.Rows.Count > 0)
                        {
                            BotonExportar.Enabled = true;
                            BotonExcel.Enabled = true;
                        }
                        else
                        {
                            BotonExportar.Enabled = false;
                            BotonExcel.Enabled = false;
                        } 

					}else Mensaje.Error("Debe colocar el nro de los meses desde y hasta para realizar la consulta","Reportes Profit");
					break;
				case 8: // Reporte Morosos + Carta Legal
					if(textMesHasta.Text!="" && textMesDesde.Text!="")
					{
						ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

						dt = ObjProfit.Reporte_Morosos(Convert.ToInt32(textMesDesde.Text),Convert.ToInt32(textMesHasta.Text));

						GridResultado.SetDataBinding(dt,"");

                        if (dt.Rows.Count > 0)
                        {
                            BotonExportar.Enabled = true;
                            BotonExcel.Enabled = true;
                        }
                        else
                        {
                            BotonExportar.Enabled = false;
                            BotonExcel.Enabled = false;
                        } 
				
					}else Mensaje.Error("Debe colocar el nro de los meses desde y hasta para realizar la consulta","Reportes Profit");

					break;
				case 9: // Reporte Listado Todos Los Morosos
						ObjProfit = new Profit.Reportes.Profit(Fecha.Value, rbconexion1.Checked);

						dt = ObjProfit.Reporte_Morosos_Todos();

						GridResultado.SetDataBinding(dt,"");

                        if (dt.Rows.Count > 0)
                        {
                            BotonExportar.Enabled = true;
                            BotonExcel.Enabled = true;
                        }
                        else
                        {
                            BotonExportar.Enabled = false;
                            BotonExcel.Enabled = false;
                        } 
				
					break;
			}
			labelGenerando.Visible = false;
		}

		private void BotonExportar_Click(object sender, System.EventArgs e)
		{
			SaveFileDialog SFD = new SaveFileDialog();
            SFD.Filter = "Archivo Text (*.txt)|*.txt|Todos los Archivos (*.*)|*.*";
            SFD.FilterIndex = 1;

			switch(ComboReporte.SelectedIndex)
			{
				case 0:
					//Cuentas Por Cobrar

					SFD.ShowDialog(this);
					SFD.AddExtension = true;
					SFD.Filter = "Archivos de Text .TXT|TXT";
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{
							Saint.Reportes.ReporteCxC RptCXC = new Saint.Reportes.ReporteCxC();

							RptCXC.GenerarReporte(ds,Fecha.Value); 

							RptCXC.ExportarReporte(SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);

						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					
					break;
				case 1:
					//Experiencia
					
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{
							Saint.Reportes.Saint.ExportarExperiencia(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 2:
					//Experiencia
					
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{
							Profit.Reportes.Profit.ExportarExperiencia(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 3:
				case 4:
				case 5:
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{

							ds = new DataSet();
							ds.Tables.Add(dt);

							Profit.Reportes.Carta_2y3 rpt = new GrupoEmporium.Profit.Reportes.Carta_2y3();

							rpt.GenerarReporte(ds,DateTime.Now);
							rpt.ExportarReporte(SFD.FileNames[0]);

							//Profit.Reportes.Profit.ExportarCartas(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 6:
					SFD.ShowDialog(this);
					//SFD.AddExtension = true;
					//SFD.Filter = "PDF|*.pdf";
					//SFD.DefaultExt = "*.pdf";

					if(SFD.FileName.Trim()!="")
					{
						try
						{

							ds = new DataSet();
							ds.Tables.Add(dt);

							Profit.Reportes.Carta_Encuesta rpt = new GrupoEmporium.Profit.Reportes.Carta_Encuesta();

							rpt.GenerarReporte(ds,DateTime.Now);
							rpt.ExportarReporte(SFD.FileNames[0]);

							//Profit.Reportes.Profit.ExportarCartas(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 7:
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{

							ds = new DataSet();
							ds.Tables.Add(dt);

							Profit.Reportes.UltimoAviso rpt = new GrupoEmporium.Profit.Reportes.UltimoAviso();

							rpt.GenerarReporte(ds,DateTime.Now);
							rpt.ExportarReporte(SFD.FileNames[0]);

							//Profit.Reportes.Profit.ExportarCartas(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 8:
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{

							ds = new DataSet();
							ds.Tables.Add(dt);

							Profit.Reportes.CartaLegal rpt = new GrupoEmporium.Profit.Reportes.CartaLegal();

							rpt.GenerarReporte(ds,DateTime.Now);
							rpt.ExportarReporte(SFD.FileNames[0]);

							//Profit.Reportes.Profit.ExportarCartas(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				case 9:
					//Experiencia
					
					SFD.ShowDialog(this);
					if(SFD.FileNames[0].Trim()!="")
					{
						try
						{
							Profit.Reportes.Profit.ExportarMorosos(dt,SFD.FileNames[0]);
							Mensaje.Informar("Reporte Generado",this);
						}
						catch(Exception ex)
						{Mensaje.Error(ex.Message,this);}
					}
					break;
				default:
					
					break;
			}
		
		}

		private void BtnMigrar_Click(object sender, System.EventArgs e)
		{
			FormMigrar Fmigrar = new FormMigrar();
			Fmigrar.ShowDialog(this);

		}

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void BotonExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog SFD = new SaveFileDialog();
            SFD.Filter = "Archivos Excel (*.xls)|*.xls|Todos los Archivos (*.*)|*.*";
            SFD.FilterIndex = 1;

            switch (ComboReporte.SelectedIndex)
            {
                case 2:
                    //Experiencia
                    SFD.ShowDialog(this);
                    if (SFD.FileNames[0].Trim() != "")
                    {
                        try
                        {
                            Profit.Reportes.Profit.ExportarExperienciaExcel(dt, SFD.FileNames[0]);
                            Mensaje.Informar("Reporte Generado", this);
                        }
                        catch (Exception ex)
                        { Mensaje.Error(ex.Message, this); }
                    }
                    break;
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                    //Cartas
                    SFD.ShowDialog(this);
                    if (SFD.FileNames[0].Trim() != "")
                    {
                        try
                        {
                            Profit.Reportes.Profit.ExportarCartasExcel(dt, SFD.FileNames[0]);
                            Mensaje.Informar("Reporte Generado", this);
                        }
                        catch (Exception ex)
                        { Mensaje.Error(ex.Message, this); }
                    }
                    break;
                case 9:
                    //Listado Morosos
                    SFD.ShowDialog(this);
                    if (SFD.FileNames[0].Trim() != "")
                    {
                        try
                        {
                            Profit.Reportes.Profit.ExportarMorososExcel(dt, SFD.FileNames[0]);
                            Mensaje.Informar("Reporte Generado", this);
                        }
                        catch (Exception ex)
                        { Mensaje.Error(ex.Message, this); }
                    }
                    break;
            }
        }
	}
}
