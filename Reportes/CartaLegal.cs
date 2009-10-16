using System;
using System.Collections;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using GrupoEmporium.Datos;
using GrupoEmporium.Saint.Reportes;
using GrupoEmporium.Reportes.PDF;

namespace GrupoEmporium.Profit.Reportes
{
	/// <summary>
	/// Clase base para generar reportes PDF
	/// </summary>
	public class CartaLegal:ReportePDF
	{
		DateTime Fecha;
		int i= 0;

		public Font F_TituloB = new Font("Arial",12,FontStyle.Bold);
		public Font F_OtrosB = new Font("Times New Roman",12,FontStyle.Regular);

		Font LetrasTabla = new Font(FontFamily.GenericSansSerif,8,FontStyle.Regular);

		DataSet DSCxC;

		#region Constructor

		public CartaLegal(System.Windows.Forms.DataGridTableStyle DTEstilo)
		{
			EstiloTabla = DTEstilo;
		}
		public CartaLegal()
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

				AnchoMiTablaPDF[0] = 10;
				AnchoMiTablaPDF[1] = 10;
				AnchoMiTablaPDF[2] = 10;
				AnchoMiTablaPDF[3] = 10;
				AnchoMiTablaPDF[4] = 10;
				AnchoMiTablaPDF[5] = 10;
				AnchoMiTablaPDF[6] = 20;
				AnchoMiTablaPDF[7] = 20;
/*				AnchoMiTablaPDF[8] = 10;
				AnchoMiTablaPDF[9] = 10;
				AnchoMiTablaPDF[10] = 10;
				AnchoMiTablaPDF[11] = 10;
*/
				MiTablaPDF.SetColumnsWidth(AnchoMiTablaPDF);

				// Modificamos la alineacion de las columnas
				//MiTablaPDF.Columns[0].SetContentAlignment(ContentAlignment.MiddleRight);
				//MiTablaPDF.Columns[1].SetContentAlignment(ContentAlignment.MiddleLeft);

				// Cambiamos los porcentajes de las areas
				CambiarPorcentajeAreas(1,20,60,18,1);

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

				for (i=0;i<DSCxC.Tables[0].Rows.Count;i++)
				{
                    if (DSCxC.Tables[0].Rows[i]["seleccion"].ToString() == "True")
                    {

                        Pagina++;
                        PdfPage NuevaPaginaPDF = DPDF.NewPage();
                        //PdfTablePage NuevaTablaPaginaPDF=TPDF.CreateTablePage(Detalle);

                        if (Inicial) // primera pagina
                        {
                            ConstruirEncabezadoInforme(ref NuevaPaginaPDF);
                            //						ConstruirEncabezadoPagina(ref NuevaPaginaPDF);
                        }

                        //ConstruirDetalle(ref NuevaTablaPaginaPDF,ref NuevaPaginaPDF);
                        ConstruirEncabezadoPagina(ref NuevaPaginaPDF);
                        ConstruirDetalle(ref NuevaPaginaPDF);
                        ConstruirPieInfome(ref NuevaPaginaPDF);

                        if (TPDF.AllTablePagesCreated) ConstruirPiePagina(ref NuevaPaginaPDF);

                        NuevaPaginaPDF.SaveToDocument();

                        if (Inicial)
                        {
                            //CambiarPorcentajeAreas(1,1,TamDetalle+TamEncabezadoInfome+TamEncabezadoPagina-2,TamPieInfome,TamPiePagina);
                            //GenerarPorcentajeAreas(ref DPDF);
                            //GenerarAreasBasicas(ref DPDF);
                            Inicial = false;
                        }
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
//			PdfArea AreatextPDF_A = EncabezadoInfome.InnerAreaP(0,8);
//			PdfArea AreatextPDF_B = EncabezadoInfome.InnerAreaP(0,20);
//			PdfArea AreatextPDF_C = EncabezadoInfome.InnerAreaP(0,32);
//
//			PdfTextArea textPDF_A = new PdfTextArea(F_TituloB,AreatextPDF_A,
//			"Barquisimeto, " + DateTime.Now.ToString("dd/MM/yy"),ContentAlignment.MiddleLeft);
//
//			PdfTextArea textPDF_B = new PdfTextArea(F_TituloB,AreatextPDF_B,
//				"DE: Agencia Royal 33, C.A. RIF: J-00000525-7",ContentAlignment.TopLeft);

//			PdfTextArea textPDF_C = new PdfTextArea(F_TituloB,AreatextPDF_C,
//				"PARA: ",ContentAlignment.TopLeft);

			//
			//PPDF.Add(MiImagenPDF,EncabezadoInfome,200);
//			PPDF.Add(textPDF_A);
//			PPDF.Add(textPDF_B);
			#endregion
		}

		protected override void ConstruirEncabezadoPagina(ref PdfPage PPDF)
		{
			#region Ejemplo
			PdfArea AreatextPDF_A = EncabezadoPagina.InnerAreaP(5,5);
			PdfArea AreatextPDF_B = EncabezadoPagina.InnerAreaP(5,30);
			PdfArea AreatextPDF_C = EncabezadoPagina.InnerAreaP(5,45);
			PdfArea AreatextPDF_D = EncabezadoPagina.InnerAreaP(5,60);
			PdfArea AreatextPDF_E = EncabezadoPagina.InnerAreaP(5,75);
			PdfArea AreatextPDF_F = EncabezadoPagina.InnerAreaP(5,90);
			PdfArea AreatextPDF_G = EncabezadoPagina.InnerAreaP(5,105);

			//
			PdfTextArea textPDF_A = new PdfTextArea(F_TituloB,AreatextPDF_A,
				"CARTA LEGAL",ContentAlignment.TopCenter);
			//
			PdfTextArea textPDF_B = new PdfTextArea(F_TituloB,AreatextPDF_A,
				"Barquisimeto, " + DateTime.Now.ToString("dd/MM/yy"),ContentAlignment.TopLeft);
			//
			PdfTextArea textPDF_C = new PdfTextArea(F_TituloB,AreatextPDF_B,
				"DE: Agencia Royal 33, C.A. RIF: J-00000525-7",ContentAlignment.TopLeft);

			PdfTextArea textPDF_D = new PdfTextArea(F_TituloB,AreatextPDF_C,
				"PARA: " + DSCxC.Tables[0].Rows[i]["cliente"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_E = new PdfTextArea(F_TituloB,AreatextPDF_D,
				"C.I.: " + DSCxC.Tables[0].Rows[i]["cedula"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_F = new PdfTextArea(F_TituloB,AreatextPDF_E,
				"DIRECCIÓN: " + DSCxC.Tables[0].Rows[i]["direccion"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_G = new PdfTextArea(F_TituloB,AreatextPDF_F,
				"TELEFONO: " + DSCxC.Tables[0].Rows[i]["telefono"].ToString(),ContentAlignment.TopLeft);

			//
			//PPDF.Add(MiImagenPDF,EncabezadoInfome,200);
			PPDF.Add(textPDF_A);
			PPDF.Add(textPDF_B);
			PPDF.Add(textPDF_C);
			PPDF.Add(textPDF_D);
			PPDF.Add(textPDF_E);
			PPDF.Add(textPDF_F);
			PPDF.Add(textPDF_G);
			#endregion

		}

		protected void ConstruirDetalle(ref PdfPage PPDF)
		{
			#region Ejemplo
//			PPDF.Add(TPPDF);

			PdfArea AreatextPDF_A = Detalle.InnerAreaP(5,0);
			PdfArea AreatextPDF_B = Detalle.InnerAreaP(5,20);
			PdfArea AreatextPDF_C = Detalle.InnerAreaP(5,60);
			PdfArea AreatextPDF_D = Detalle.InnerAreaP(5,80);
			PdfArea AreatextPDF_E = Detalle.InnerAreaP(5,100);
			PdfArea AreatextPDF_F = Detalle.InnerAreaP(5,120);
			PdfArea AreatextPDF_G = Detalle.InnerAreaP(5,140);
			PdfArea AreatextPDF_H = Detalle.InnerAreaP(5,140);


			string Carta_P1 =	"    En vista de haber hecho caso omiso a nuestro ULTIMO AVISO DE COBRO le participamos " +
								"que a partir de la presente fecha su caso está siendo tratado por el Departamento Legal " + 
								"de la empresa y se ha autorizado el retiro de la mercancía; en caso de no tener respuesta " +
								"satisfactoria de su parte en las próximas 24 horas.";

			string Carta_P2 =	"    Permitanos recordarle que aún es tiempo de alcanzar un arreglo para resolver su problema";

			string Carta_P3 =	"    Comuníquese urgentemente con el Dpto. Legal de la empresa a los Telefonos: 2329960 - 2326872.";


			PdfTextArea textPDF_A = new PdfTextArea(F_OtrosB,AreatextPDF_A,
				"CARTA LEGAL",ContentAlignment.TopCenter);

			PdfTextArea textPDF_B = new PdfTextArea(F_OtrosB,AreatextPDF_B,
				Carta_P1,ContentAlignment.TopLeft);

			PdfTextArea textPDF_C = new PdfTextArea(F_OtrosB,AreatextPDF_C,
				Carta_P2,ContentAlignment.TopLeft);

			PdfTextArea textPDF_D = new PdfTextArea(F_OtrosB,AreatextPDF_D,
				Carta_P3,ContentAlignment.TopLeft);

			
			//PPDF.Add(textPDF_A);
			PPDF.Add(textPDF_B);
			PPDF.Add(textPDF_C);
			PPDF.Add(textPDF_D);
			//PPDF.Add(textPDF_E);
			//PPDF.Add(textPDF_F);
			//PPDF.Add(textPDF_G);
			//PPDF.Add(textPDF_H);
			#endregion
		}

		protected override void ConstruirPieInfome(ref PdfPage PPDF)
		{

			#region Ejemplo
			PdfArea AreatextPDF_A = PieInfome.InnerAreaP(5,1);
			PdfArea AreatextPDF_B = PieInfome.InnerAreaP(5,15);
			PdfArea AreatextPDF_C = PieInfome.InnerAreaP(5,30);
			PdfArea AreatextPDF_D = PieInfome.InnerAreaP(5,45);
			PdfArea AreatextPDF_F = PieInfome.InnerAreaP(5,60);
			PdfArea AreatextPDF_E = PieInfome.InnerAreaP(5,75);

			//
			PdfTextArea textPDF_A = new PdfTextArea(F_TituloB,AreatextPDF_A,
				"Barquisimeto, " + DateTime.Now.ToString("dd/MM/yy"),ContentAlignment.TopLeft);
			//
			PdfTextArea textPDF_B = new PdfTextArea(F_TituloB,AreatextPDF_B,
				"DE: Agencia Royal 33, C.A. RIF: J-00000525-7",ContentAlignment.TopLeft);

			PdfTextArea textPDF_C = new PdfTextArea(F_TituloB,AreatextPDF_C,
				"PARA: " + DSCxC.Tables[0].Rows[i]["cliente"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_D = new PdfTextArea(F_TituloB,AreatextPDF_D,
				"C.I.: " + DSCxC.Tables[0].Rows[i]["cedula"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_E = new PdfTextArea(F_TituloB,AreatextPDF_E,
				"DIRECCIÓN: " + DSCxC.Tables[0].Rows[i]["direccion"].ToString(),ContentAlignment.TopLeft);

			PdfTextArea textPDF_F = new PdfTextArea(F_TituloB,AreatextPDF_F,
				"TELEFONO: " + DSCxC.Tables[0].Rows[i]["telefono"].ToString(),ContentAlignment.TopLeft);

			//
			//PPDF.Add(MiImagenPDF,EncabezadoInfome,200);
			//PPDF.Add(textPDF_A);
			//PPDF.Add(textPDF_B);
			//PPDF.Add(textPDF_C);
			//PPDF.Add(textPDF_D);
			//PPDF.Add(textPDF_E);
			//PPDF.Add(textPDF_F);
			#endregion
		}

		protected override void ConstruirPiePagina(ref PdfPage PPDF)
		{
			#region Ejemplo
			PdfArea AreatextPDF_NroPag = PiePagina.InnerAreaP(PiePagina.Width - 80,0);

			PdfTextArea textPDF_NroPag = new PdfTextArea(F_EncabezadoA,AreatextPDF_NroPag,
				"Página " + Pagina.ToString());

//			PPDF.Add(textPDF_NroPag);
			#endregion
		}

		#endregion

		#endregion

	}
}
