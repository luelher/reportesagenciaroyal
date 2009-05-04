using System;
using System.Collections;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using GrupoEmporium.Datos;
using GrupoEmporium.Saint.Reportes;
using GrupoEmporium.Reportes.PDF;

namespace GrupoEmporium.Saint.Reportes
{
	/// <summary>
	/// Clase base para generar reportes PDF
	/// </summary>
	public class ReporteCxC:ReportePDF
	{
		DateTime Fecha;

		Font LetrasTabla = new Font(FontFamily.GenericSansSerif,8,FontStyle.Regular);

		DataSet DSCxC;

		#region Constructor

		public ReporteCxC(System.Windows.Forms.DataGridTableStyle DTEstilo)
		{
			EstiloTabla = DTEstilo;
		}
		public ReporteCxC()
		{
		}

		#endregion

		#region Metodos A Reemplazar en las clases derivada

		#region Generar Reporte

		public bool GenerarReporte(DataSet DS,DateTime F)
		{
			#region Ejemplo
			try
			{
				Fecha = F;

				DSCxC = DS;

				// Creando el Documento Pdf Base
				MiDocumentoPDF = new PdfDocument(MiFormatoDocumento);

				// Calculamos las filas y columnas y lo guardamos en Filas,Columnas,PCortes
				CalcularPuntosCorte(DS.Tables[0]);
				
				// Creando Tabla de datos
				MiTablaPDF=MiDocumentoPDF.NewTable(F_TablaA,Filas,Columnas,PCortes);

				// Llenando la TablaPDF con los datos del DT
				MiTablaPDF.ImportDataTable(DS.Tables[0]);

				// Configurando el Formato de fecha
				//MiTablaPDF.Columns[7].SetContentFormat("{0:dd/MM/yyyy}");

				// Colocando el estilo a la Tabla PDF
				ConfigurarEstiloTablaPDF(ref MiTablaPDF);
				MiTablaPDF.SetFont(LetrasTabla);

				// Organizando el ancho de las columnas
				int[] AnchoMiTablaPDF = AnchoTablaPDF(ref MiTablaPDF);

//				 "Cedula" ("System.String"));
//				 "Nombre" ("System.String"));
//				 "Cuotas" ("System.Int32"));
//				 "Total" ("System.Double"));
//				 "Vencidas" ("System.Int32"));					
//				 "TotalVencido" ("System.Double"));
//				 "PorVencer" ("System.Int32"));
//				 "Total0a30" ("System.Double"));
//				 "Total30a60" ("System.Double"));
//				 "Total60a90" ("System.Double"));
//				 "Total90a120" ("System.Double"));
//				 "Mayor120" ("System.Double"));

				AnchoMiTablaPDF[0] = 5;
				AnchoMiTablaPDF[1] = 10;
				AnchoMiTablaPDF[2] = 5;
				AnchoMiTablaPDF[3] = 10;
				AnchoMiTablaPDF[4] = 5;
				AnchoMiTablaPDF[5] = 10;
				AnchoMiTablaPDF[6] = 5;
				AnchoMiTablaPDF[7] = 10;
				AnchoMiTablaPDF[8] = 10;
				AnchoMiTablaPDF[9] = 10;
				AnchoMiTablaPDF[10] = 10;
				AnchoMiTablaPDF[11] = 10;
				MiTablaPDF.SetColumnsWidth(AnchoMiTablaPDF);

				// Modificamos la alineacion de las columnas
				//MiTablaPDF.Columns[0].SetContentAlignment(ContentAlignment.MiddleRight);
				//MiTablaPDF.Columns[1].SetContentAlignment(ContentAlignment.MiddleLeft);

				// Cambiamos los porcentajes de las areas
				CambiarPorcentajeAreas(10,5,77,6,2);

//				MiImagenPDF = MiDocumentoPDF.NewImage(@"Imagenes\Logo.jpg");

				ConstruirPDF(ref MiDocumentoPDF,ref MiTablaPDF);

				return true;

			}
			catch(Exception ex)
			{
				string err = ex.Message;
				return false;
			}
			#endregion
		}

		#endregion

		#region Metodos para colocar Objetos en áreas básicas

		protected override bool ConstruirPDF(ref PdfDocument DPDF,ref PdfTable TPDF)
		{
			#region Ejemplo
			try
			{
				bool Inicial=true;
				GenerarPorcentajeAreas(ref DPDF);
				GenerarAreasBasicas(ref DPDF);

				while (!TPDF.AllTablePagesCreated)
				{
					Pagina++;
					PdfPage NuevaPaginaPDF=DPDF.NewPage();
					PdfTablePage NuevaTablaPaginaPDF=TPDF.CreateTablePage(Detalle);

					if(Inicial) // primera pagina
					{
						ConstruirEncabezadoInforme(ref NuevaPaginaPDF);
						ConstruirEncabezadoPagina(ref NuevaPaginaPDF);
					}

					ConstruirDetalle(ref NuevaTablaPaginaPDF,ref NuevaPaginaPDF);
					ConstruirPieInfome(ref NuevaPaginaPDF);

					if(TPDF.AllTablePagesCreated) ConstruirPiePagina(ref NuevaPaginaPDF);

					NuevaPaginaPDF.SaveToDocument();

					if(Inicial)
					{
						CambiarPorcentajeAreas(1,1,TamDetalle+TamEncabezadoInfome+TamEncabezadoPagina-2,TamPieInfome,TamPiePagina);
						GenerarPorcentajeAreas(ref DPDF);
						GenerarAreasBasicas(ref DPDF);
						Inicial=false;
					}

				}
				return true;
			}
			catch(Exception ex)
			{
				string err = ex.Message;
				return false;
			}
			#endregion
		}

		protected override void ConstruirEncabezadoInforme(ref PdfPage PPDF)
		{
			#region Ejemplo
			PdfArea AreatextPDF_A = EncabezadoInfome.InnerAreaP(0,8);
			PdfArea AreatextPDF_B = EncabezadoInfome.InnerAreaP(0,EncabezadoInfome.Height-20);
//
			PdfTextArea textPDF_A = new PdfTextArea(F_TituloA,AreatextPDF_A,
			"Agencia Royal Victor Marmol C.A.",ContentAlignment.TopCenter);
//
			PdfTextArea textPDF_B = new PdfTextArea(F_TituloB,AreatextPDF_B,
				"Informe Cuentas por Cobrar",ContentAlignment.TopCenter);
//
//			PPDF.Add(MiImagenPDF,EncabezadoInfome,200);
			PPDF.Add(textPDF_A);
			PPDF.Add(textPDF_B);
			#endregion
		}

		protected override void ConstruirEncabezadoPagina(ref PdfPage PPDF)
		{
			#region Ejemplo
//			Opciones Opc = new Opciones();
//
//			//EncabezadoPagina.ToRectangle(Color.Blue,10,Color.CornflowerBlue);
			PdfArea AreatextPDF_A = EncabezadoPagina.InnerAreaP(0,0);
//			PdfArea AreatextPDF_CodSeccion = EncabezadoPagina.InnerAreaP(0,15);
//			PdfArea AreatextPDF_Materia = EncabezadoPagina.InnerAreaP(100,15);
//			PdfArea AreatextPDF_Profesor = EncabezadoPagina.InnerAreaP(0,30);
//			PdfArea AreatextPDF_Planificado = EncabezadoPagina.InnerAreaP(0,45);
//
			PdfTextArea textPDF_Fecha = new PdfTextArea(F_EncabezadoA,AreatextPDF_A,
				"Hasta: " + Fecha.ToString("dd/MM/yyyy"));
//
//			PdfTextArea textPDF_CodSeccion = new PdfTextArea(F_EncabezadoA,AreatextPDF_CodSeccion,
//				"Código Sección: " + Secc.CodSeccion);
//
//			PdfTextArea textPDF_Materia = new PdfTextArea(F_EncabezadoA,AreatextPDF_Materia,
//				"Materia: " + Secc.NombreMateria);
//
//			PdfTextArea textPDF_Profesor = new PdfTextArea(F_EncabezadoA,AreatextPDF_Profesor,
//				"Profesor: " + Secc.NombreProfesor);
//
//			PdfTextArea textPDF_Planificado;
//			if(Secc.Planificado) 
//				textPDF_Planificado = new PdfTextArea(F_EncabezadoA,AreatextPDF_Planificado,
//				"Planificado: Si" );
//			else textPDF_Planificado = new PdfTextArea(F_EncabezadoA,AreatextPDF_Planificado,
//				"Planificado: No" );
//
			PPDF.Add(textPDF_Fecha);
//			PPDF.Add(textPDF_Semestre);
//			PPDF.Add(textPDF_Materia);
//			PPDF.Add(textPDF_Profesor);
//			PPDF.Add(textPDF_Planificado);
//			//PPDF.Add(EncabezadoPagina.ToRectangle(Color.Blue,1,Color.CornflowerBlue));
			#endregion
		}

		protected override void ConstruirDetalle(ref PdfTablePage TPPDF, ref PdfPage PPDF)
		{
			#region Ejemplo
			PPDF.Add(TPPDF);
			#endregion
		}

		protected override void ConstruirPieInfome(ref PdfPage PPDF)
		{

			#region Ejemplo
			PdfArea AreatextPDF_Totales = PieInfome.InnerAreaP(0,5);
			PdfArea AreatextPDF_TotalPorCobrar = PieInfome.InnerAreaP(PieInfome.Width - 150,5);
//			PdfArea AreatextPDF_Linea = PieInfome.InnerAreaP(0,0);
//
			PdfTextArea textPDF_Totales = new PdfTextArea(F_EncabezadoA,AreatextPDF_Totales,
				"TOTALES -----> ");

			PdfTextArea TextPDF_TotalPorCobrar = new PdfTextArea(F_EncabezadoB,AreatextPDF_TotalPorCobrar,
				"Total Por Cobrar: " + DSCxC.Tables[1].Rows[0]["Total"].ToString() + "%");
//
//			PdfLine lineaPDF = AreatextPDF_Linea.UpperBound(Color.Black,2);
//
//			PPDF.Add(lineaPDF);
//			PPDF.Add(textPDF_Totales);
//			PPDF.Add(textPDF_Evaluado);
			#endregion
		}

		protected override void ConstruirPiePagina(ref PdfPage PPDF)
		{
			#region Ejemplo
			PdfArea AreatextPDF_NroPag = PiePagina.InnerAreaP(PiePagina.Width - 80,0);

			PdfTextArea textPDF_NroPag = new PdfTextArea(F_EncabezadoA,AreatextPDF_NroPag,
				"Página " + Pagina.ToString());

			PPDF.Add(textPDF_NroPag);
			#endregion
		}

		#endregion

		#endregion

	}
}
