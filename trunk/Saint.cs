using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Diagnostics;
using System.Windows.Forms;
using GrupoEmporium.Reportes.PDF;
using GrupoEmporium.Datos;
using GrupoEmporium.Varias;
using GrupoEmporium.Mensajes;

namespace GrupoEmporium.Saint.Reportes
{
	/// <summary>
	/// Clase para generar reportes de Saint Enterprise.
	/// </summary>
	public class Saint
	{
		#region Variables

		string SQL="";
		DateTime _LaFecha;

		GrupoEmporium.Varias.ClaseDocumentosXML configxml;

		clsBDConexion strConexion;
		string cadenaconexion;
		clsBDConexion strConexion_Profit;
		string cadenaconexion_Profit;
		
		//Este objeto conexion debe ser reemplazado al cambiar a SQL Server
		public OleDbConnection Conexion;
		public OleDbConnection Conexion_Profit;


		private static string Tab="	";

		public struct ResumenCxC
		{
			public string IDCliente;
			public string Cliente;
			public string FechaC;

			public double Vencido;
			public int CuotasVencidas; 
			public int Cuotas; 
			public int CuotasxVencer;

			public double xVencer0_30;
			public double xVencer30_60;
			public double xVencer60_90;
			public double xVencer90_120;
			public double xVencerMayor120;
			public double Total;

			public ResumenCxC(string strIDCliente)
			{
				IDCliente=strIDCliente;
				Cliente="";
				FechaC = "";

				Vencido=0.0;
				CuotasVencidas=0; 
				Cuotas=0; 
				CuotasxVencer=0;

				xVencer0_30=0.0;
				xVencer30_60=0.0;
				xVencer60_90=0.0;
				xVencer90_120=0.0;
				xVencerMayor120=0.0;
				Total=0.0;
			}

		}


		private struct TotalesCxC
		{
			public double TotalVencido;
			public int TotalCuotasVencidas; 
			public int TotalCuotasxVencer;

			public double TotalxVencer0_30;
			public double TotalxVencer30_60;
			public double TotalxVencer60_90;
			public double TotalxVencer90_120;
			public double TotalxVencerMayor120;
			public double TotalxC;

			public TotalesCxC(int c)
			{
				c=0;
				TotalVencido=0.0;
				TotalCuotasVencidas=c; 
				TotalCuotasxVencer=c;

				TotalxVencer0_30=0.0;
				TotalxVencer30_60=0.0;
				TotalxVencer60_90=0.0;
				TotalxVencer90_120=0.0;
				TotalxVencerMayor120=0.0;
				TotalxC=0.0;
			}

		}


		public struct DetalleFactura
		{
			public string CodItem;
			public string Item;
			public string Descripcion;
			public double Costo;

			public DetalleFactura(int c)
			{
				CodItem="";
				Item="";
				Descripcion="";
				Costo=0.0;
			}

		}



		public struct Factura
		{
			public string IDCliente;
			public string Cliente;
			public string Direccion;
			public string Telefono;

			public string NroFactura;
			public DateTime FechaE;
			public double MontoTotal;
			public double Impuesto;
			public double MontoNeto;
			public double Adelanto;
			public double Intereses;
			public double PagoMensual;
			public int Giros;
			public DateTime FechaCancelacion;
			private byte _Experiencia;

			public Factura(string strNroFactura)
			{
				IDCliente="";
				Cliente="";

				Telefono="";

				NroFactura=strNroFactura;
				FechaE=DateTime.MinValue;
				MontoTotal=0.0;
				PagoMensual=0.0;
				Giros=0;
				FechaCancelacion=DateTime.MinValue;
				Direccion="";
				_Experiencia=0;
				MontoNeto = 0.0;
				Adelanto = 0.0;
				Intereses = 0.0;
				Impuesto = 0.0;

			}


			public byte Experiencia
			{
				get {return _Experiencia;}
				set {if (value > _Experiencia) _Experiencia=value;}
			}

		}


		public struct CxC
		{
			public DateTime FechaE;
			public DateTime FechaV;
			public double Monto;
			public double Saldo;
			public int NroCuota;
			public bool Cancelada;

			public CxC(int c)
			{
				FechaE = DateTime.MinValue;
				FechaV = DateTime.MinValue;
				Monto = 0.0;
				Saldo = 0.0;
				NroCuota = 0;
				Cancelada = false;
			}

 
		}

		/// <summary>
		/// Contiene la informacion de cuotas y la transacción de credito del cliente.
		/// </summary>
		public struct Documentos
		{
			public Factura fact;
			public DetalleFactura[] detfact;
			public CxC[] cxc;
		}

		#endregion

		#region Contructor
		public Saint()
		{
			CargarConfig();

			_LaFecha = DateTime.Now;

			try
			{
				Conexion.Open();
				Mensajes.Mensaje.Error(Conexion.State.ToString(),"Saint Reportes");
			}
			catch(Exception ex)
			{
				Mensajes.Mensaje.Error(ex.Message,"Saint Reportes");
			}

		}
		public Saint(DateTime Fecha)
		{

			CargarConfig();

			_LaFecha = Fecha;

			try
			{
				Conexion.Open();
				Mensajes.Mensaje.Error(Conexion.State.ToString(),"Saint Reportes");
			}
			catch(Exception ex)
			{
				Mensajes.Mensaje.Error(ex.Message,"Saint Reportes");
			}

		}
		#endregion

		#region Propiedades

		public DateTime LaFecha
		{
			get{return _LaFecha;}
			set{_LaFecha=value;}
		}


		#endregion

		#region Metodos Públicos

		public DataTable Reporte_Experiencia()
		{
			if(Conexion.State == ConnectionState.Open)
			{

				DataTable dt = new DataTable();
				DataTable dtReporte = new DataTable();
				DataRow dr;
			
				#region SQL Union Facturas

				SQL =	" SELECT " +
					" FacturasViejas.Cedula AS CodClie, " +
					" SACLIE.Descrip, " +
					" SACLIE.Direc1, " +
					" SACLIE.Direc2, " +
					" SACLIE.Telef, " +
					" cast(FacturasViejas.NFactura as varchar(10)) AS NumeroD, " +
					" FacturasViejas.GNumero AS NGiros, " +
					" FacturasViejas.FechaE, " +
					" FacturasViejas.FechaV, " +
					" CAST(FacturasViejas.MontoV AS decimal(13)) AS MtoFinanc " +
					" FROM " +
					" FacturasViejas INNER JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) " +
					" ON (FacturasViejas.Cedula=SACLIE.CodClie) AND ((FacturasViejas.FechaE) = (SAACXC.FechaE))  " +
					" WHERE " +
					" DateDiff(dd,FacturasViejas.FechaE,CAST('" + _LaFecha.ToString("yyMMdd") + "' AS DATETIME)) > 0  AND  " +
					" FacturasViejas.GNumero<>'0'  AND  " +
					" FacturasViejas.MontoV>0  " +
					" GROUP BY " +
					" FacturasViejas.Cedula, " +
					" FacturasViejas.NFactura, " +
					" FacturasViejas.GNumero, " +
					" FacturasViejas.FechaE, " +
					" FacturasViejas.FechaV, " +
					" FacturasViejas.MontoV, " +
					" SACLIE.Descrip, " +
					" SACLIE.Direc1, " +
					" SACLIE.Direc2, " +
					" SACLIE.Telef " +
					" UNION " +
					" SELECT " +
					" SAFACT.CodClie, " +
					" SACLIE.Descrip, " +
					" SACLIE.Direc1, " +
					" SACLIE.Direc2, " +
					" SACLIE.Telef, " +
					" SAFACT.NumeroD, " +
					" SAFACT.NGiros, " +
					" SAFACT.FechaE, " +
					" SAFACT.FechaV, " +
					" SAFACT.MtoFinanc  " +
					" FROM " +
					" SAFACT INNER JOIN " +
					" (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) ON " +
					" ((SAFACT.CodClie=SACLIE.CodClie) AND DATEDIFF(dd,SAFACT.FechaE,SAACXC.FechaE)=0) " +
					" WHERE " +
					" DATEDIFF(dd,SAFACT.FechaE,CAST('" + _LaFecha.ToString("yyMMdd") + "' AS DATETIME)) > 0 AND  " +
					" SAFACT.NGiros>0  AND " +
					" SAFACT.MtoFinanc >0 " +
					" GROUP BY " +
					" SAFACT.CodClie, " +
					" SAFACT.NumeroD, " +
					" SAFACT.NGiros, " +
					" SAFACT.FechaE, " +
					" SAFACT.FechaV, " +
					" SAFACT.MtoFinanc, " +
					" SACLIE.Descrip, " +
					" SACLIE.Direc1, " +
					" SACLIE.Direc2, " +
					" SACLIE.Telef  " +
					" ORDER BY " +
					" SAFACT.NumeroD; ";

				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

				Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

				#endregion

				if(dt.Rows.Count>0)
				{
					Factura RFactura;

					dtReporte.Columns.Add("Cedula",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("Nombre",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("Telefono",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("Factura",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("FechaE",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("MontoTotal",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("PagoMensual",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Giros",System.Type.GetType("System.Byte"));
					dtReporte.Columns.Add("FechaCancelacion",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("Experiencia",System.Type.GetType("System.Byte"));
				
					for(int i=0;i<dt.Rows.Count;i++)
					{
					
						RFactura = ResumenFactura(dt.Rows[i]["NumeroD"].ToString(),dt.Rows[i]["CodClie"].ToString(),(DateTime)dt.Rows[i]["FechaE"]);

						RFactura.IDCliente = dt.Rows[i]["CodClie"].ToString();
						int CantGiros = Convert.ToInt32( dt.Rows[i]["NGiros"].ToString() );
						if(CantGiros>0) RFactura.PagoMensual  = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString())/CantGiros;

						RFactura.Giros = CantGiros;
						RFactura.MontoTotal = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString());
						RFactura.FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
						//RFactura.FechaCancelacion = Convert.ToDateTime(dt.Rows[i]["FechaV"].ToString());
						RFactura.Cliente = dt.Rows[i]["Descrip"].ToString();

						if(RFactura.NroFactura !="")
						{
							dr = dtReporte.NewRow();
							dr["Cedula"] = RFactura.IDCliente;
							dr["Nombre"] = RFactura.Cliente;
							dr["Telefono"] = RFactura.Telefono;
							dr["Factura"] = RFactura.NroFactura;
							dr["FechaE"] = RFactura.FechaE;
							dr["MontoTotal"] = RFactura.MontoTotal;
							dr["PagoMensual"] = RFactura.PagoMensual;
							dr["Giros"] = RFactura.Giros;
							dr["FechaCancelacion"] = RFactura.FechaCancelacion;
							dr["Experiencia"] = RFactura.Experiencia;
							dtReporte.Rows.Add(dr);
						}
					}
					dtReporte.AcceptChanges();
					return dtReporte;
				}return new DataTable();
			}
			else return new DataTable();
		}


		public static void ExportarExperiencia(DataTable dt,string arch)
		{
			if (System.IO.File.Exists(arch)){System.IO.File.Delete(arch);}
			StreamWriter TxtFile = new StreamWriter (arch,true);

			double Mensual;
			DateTime FechaE;
			double MontoT;
			DateTime FechaV;

			for(int i=0;i<dt.Rows.Count;i++)
			{

				FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
				MontoT = Convert.ToDouble(dt.Rows[i]["MontoTotal"].ToString());
				Mensual= Convert.ToDouble(dt.Rows[i]["PagoMensual"].ToString());
				FechaV = Convert.ToDateTime(dt.Rows[i]["FechaCancelacion"].ToString());

				string Cad = dt.Rows[i]["Cedula"].ToString() + Tab +
					dt.Rows[i]["Nombre"].ToString() +Tab+
					Tab+
					dt.Rows[i]["Telefono"].ToString() +Tab+
					Tab+
					Tab+
					dt.Rows[i]["Factura"].ToString() +Tab+
					FechaE.ToString("dd/MM/yyyy")+Tab+
					MontoT.ToString("#,##0.00;($#,##0.00);0")+Tab+
					Mensual.ToString("#,##0.00;($#,##0.00);0")+Tab+
					dt.Rows[i]["Giros"].ToString() +Tab+
					FechaV.ToString("dd/MM/yyyy")+Tab+
					dt.Rows[i]["Experiencia"].ToString() +Tab;

				TxtFile.WriteLine(Cad);

			}

			TxtFile.Close();

		}


		public DataSet Reporte_Resumen_CXC()
		{

			DataSet Resultado = new DataSet();
			Resultado.DataSetName = "Resumen Cuentas Por Cobrar";

			if(Conexion.State == ConnectionState.Open)
			{
				DataTable dt = new DataTable();
				DataTable dtReporte = new DataTable("Detalles");
				DataTable dtTotalesReporte = new DataTable("Resumen");
				DataRow dr;

				TotalesCxC Totales = new TotalesCxC();
			
				SQL= "Select SAACXC.CodClie FROM SAACXC GROUP BY SAACXC.CodClie";
				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

				Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

				if(dt.Rows.Count>0)
				{
					ResumenCxC CxCCliente;

					//Falta
					//Fecha de Compra
					//Ordenardos de Menor a Mayor

					dtReporte.Columns.Add("Cedula",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("Nombre",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("Fecha Compra",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("Cuotas",System.Type.GetType("System.Int32"));
					dtReporte.Columns.Add("Total",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Vencidas",System.Type.GetType("System.Int32"));					
					dtReporte.Columns.Add("TotalVencido",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Por Vencer",System.Type.GetType("System.Int32"));
					dtReporte.Columns.Add("Total 0a30",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Total 30a60",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Total 60a90",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Total 90a120",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("Mayor 120",System.Type.GetType("System.Double"));


					dtTotalesReporte.Columns.Add("Total",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Cuotas Vencidas",System.Type.GetType("System.Int32"));					
					dtTotalesReporte.Columns.Add("Vencido",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Por Vencer",System.Type.GetType("System.Int32"));
					dtTotalesReporte.Columns.Add("Total 0a30",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Total 30a60",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Total 60a90",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Total 90a120",System.Type.GetType("System.Double"));
					dtTotalesReporte.Columns.Add("Mayor 120",System.Type.GetType("System.Double"));
				
					for(int i=0;i<dt.Rows.Count;i++)
					{
					
						CxCCliente = ResumenCxCCliente(dt.Rows[i][0].ToString());
						if(CxCCliente.IDCliente  !="")
						{
							dr = dtReporte.NewRow();
							dr["Cedula"] = CxCCliente.IDCliente;
							dr["Nombre"] = CxCCliente.Cliente;
							dr["Fecha Compra"] = Convert.ToDateTime(CxCCliente.FechaC);

							dr["Cuotas"] = CxCCliente.Cuotas;

							dr["Total"] = CxCCliente.Total;
							Totales.TotalxC += CxCCliente.Total;

							dr["Vencidas"] = CxCCliente.CuotasVencidas;
							Totales.TotalCuotasVencidas += CxCCliente.CuotasVencidas;

							dr["TotalVencido"] = CxCCliente.Vencido;
							Totales.TotalVencido += CxCCliente.Vencido;

							dr["Por Vencer"] = CxCCliente.CuotasxVencer;
							Totales.TotalCuotasxVencer += CxCCliente.CuotasxVencer;

							dr["Total 0a30"] = CxCCliente.xVencer0_30;
							Totales.TotalxVencer0_30 += CxCCliente.xVencer0_30;

							dr["Total 30a60"] = CxCCliente.xVencer30_60;
							Totales.TotalxVencer30_60 += CxCCliente.xVencer30_60;

							dr["Total 60a90"] = CxCCliente.xVencer60_90;
							Totales.TotalxVencer60_90 += CxCCliente.xVencer60_90;

							dr["Total 90a120"] = CxCCliente.xVencer90_120;
							Totales.TotalxVencer90_120 += CxCCliente.xVencer90_120;

							dr["Mayor 120"] = CxCCliente.xVencerMayor120;
							Totales.TotalxVencerMayor120  += CxCCliente.xVencer30_60;

							dtReporte.Rows.Add(dr);
							//Debug.WriteLine("Registro " + i.ToString());
						}
					}
					dtReporte.AcceptChanges();

					dr = dtTotalesReporte.NewRow();

					dr["Total"] = Totales.TotalxC;
					dr["Cuotas Vencidas"] = Totales.TotalCuotasVencidas;
					dr["Vencido"] = Totales.TotalVencido;
					dr["Por Vencer"] = Totales.TotalCuotasxVencer;
					dr["Total 0a30"] = Totales.TotalxVencer0_30;
					dr["Total 30a60"] = Totales.TotalxVencer30_60;
					dr["Total 60a90"] = Totales.TotalxVencer60_90;
					dr["Total 90a120"] = Totales.TotalxVencer90_120;
					dr["Mayor 120"] = Totales.TotalxVencerMayor120;
					dtTotalesReporte.Rows.Add(dr);

					Resultado.Tables.Add(dtReporte);
					Resultado.Tables.Add(dtTotalesReporte);
					return Resultado;
				}return new DataSet();
			} 
			else return new DataSet();



		}


		#endregion

		#region Metodos Privados

		private void CargarConfig()
		{
			
			ClaseDocumentosXML MiConfig = new ClaseDocumentosXML(@"configsaint.xml");
			if(MiConfig.Cargado)
			{
				if (MiConfig["CadenaConexion"]=="False")
				{
					if (MiConfig["Dominio"]=="True")
						strConexion = new clsBDConexion(MiConfig["DataSource"],MiConfig["InitialCatalog"]);
					else strConexion = new clsBDConexion( MiConfig["DataSource"],MiConfig["InitialCatalog"],MiConfig["Usuario"],MiConfig["Contrasena"],false);

					cadenaconexion = strConexion.StringConexion;
				}
				else
				{
					strConexion = new clsBDConexion();
					strConexion.TipoBaseDato = TipoBD.SQL_SERVER;
					strConexion.StringConexion = cadenaconexion;
					cadenaconexion = MiConfig["Conexion"];
				}

			}

			Conexion = new OleDbConnection(cadenaconexion);
			configxml = MiConfig;

			CargarConfigProfit();

		}


		private void CargarConfigProfit()
		{
			
			ClaseDocumentosXML MiConfig = new ClaseDocumentosXML(@"configsaint.xml");
			if(MiConfig.Cargado)
			{
				if (MiConfig["CadenaConexion_Profit"]=="False")
				{
					if (MiConfig["Dominio_Profit"]=="True")
						strConexion_Profit = new clsBDConexion(MiConfig["DataSource_Profit"],MiConfig["InitialCatalog_Profit"]);
					else strConexion_Profit = new clsBDConexion( MiConfig["DataSource_Profit"],MiConfig["InitialCatalog_Profit"],MiConfig["Usuario_Profit"],MiConfig["Contrasena_Profit"],false);

					cadenaconexion_Profit = strConexion_Profit.StringConexion;
				}
				else
				{
					strConexion_Profit = new clsBDConexion();
					strConexion_Profit.TipoBaseDato = TipoBD.SQL_SERVER;
					strConexion_Profit.StringConexion = cadenaconexion_Profit;
					cadenaconexion_Profit = MiConfig["Conexion_Profit"];
				}

			}

			Conexion_Profit = new OleDbConnection(cadenaconexion_Profit);
			//configxml = MiConfig;

		}


		private ResumenCxC ResumenCxCCliente(string idCliente)
		{
			DataTable dtCxC = new DataTable();
			string str="";
			ResumenCxC Resumen;

			SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
				" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
				" WHERE SAACXC.CodClie='" + idCliente + "' AND SAACXC.TipoCxc='60' " +
				" ORDER BY SAACXC.NroUnico ASC; ";
			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

			if(dtCxC.Rows.Count>0)
			{
				if(dtCxC.Rows.Count>20)
				{
					str = dtCxC.Rows[0][0].ToString();

					return new ResumenCxC("");
				}
				else
				{
					int Cuota=0;
					int Total=0;
					int Dias =0;
					double Monto=0.0;
					string strMonto;
					str = dtCxC.Rows[0]["NumeroD"].ToString();
					if(Analizar_TipoCxc(str,ref Cuota,ref Total))
					{
						Resumen = new ResumenCxC(str = dtCxC.Rows[0]["CodClie"].ToString());
						Resumen.Cliente = dtCxC.Rows[0]["Descrip"].ToString();
						Resumen.FechaC = dtCxC.Rows[0]["FechaE"].ToString().Substring(0,10);
						Resumen.Cuotas = Total;

						for(int i=0;i<dtCxC.Rows.Count;i++)
						{
							Dias = Analizar_FechaV(dtCxC.Rows[i]["FechaV"].ToString(),_LaFecha);
							strMonto = dtCxC.Rows[i]["Saldo"].ToString();
							Monto = Convert.ToDouble(strMonto);

							if(Dias==-1 && Monto!=0.0) 
							{
								Resumen.CuotasVencidas++;
								Resumen.Vencido += Monto;
							}
							else if(Dias!=-1 && Monto!=0.0)
							{
								Resumen.CuotasxVencer++;
								if (Dias>=0 && Dias<30)
									Resumen.xVencer0_30 += Monto;
								else if (Dias>=30 && Dias<60)
									Resumen.xVencer30_60 += Monto;
								else if (Dias>=60 && Dias<90)
									Resumen.xVencer60_90 += Monto;
								else if (Dias>=90 && Dias<120)
									Resumen.xVencer90_120 += Monto;
								else if(Dias>=120)
									Resumen.xVencerMayor120 += Monto;
							}
						}

						Resumen.Total = Resumen.Vencido + Resumen.xVencer0_30 + Resumen.xVencer30_60 + Resumen.xVencer60_90 + Resumen.xVencer90_120 + Resumen.xVencerMayor120;

						return Resumen;
					}
					else return new ResumenCxC("");
				}
			}
			else return new ResumenCxC("");
		}


		private Factura ResumenFactura(string idFactura,string idCliente,DateTime FechaE)
		{
			DataTable dtCxC = new DataTable();
			string str="";
			Factura Resumen;

			SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
				" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
				" WHERE SAACXC.CodClie='" + idCliente + "' AND SAACXC.TipoCxc='60' AND DATEDIFF(DD,SAACXC.FechaE, CAST('" + FechaE.ToString("yyMMdd") + "' AS DATETIME)) = 0 " +
				" ORDER BY SAACXC.NroUnico ASC; ";
			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

			Resumen = new Factura("");
			Resumen.Experiencia = 254;

			if(dtCxC.Rows.Count>0)
			{
				if(dtCxC.Rows.Count>20)
				{
					str = dtCxC.Rows[0][0].ToString();

					return Resumen;
				}
				else
				{
					int Cuota=0;
					int Total=0;
					int Dias =0;
					double Monto=0.0;
					string strMonto;
					string cedula;
					str = dtCxC.Rows[0]["NumeroD"].ToString();

					cedula = dtCxC.Rows[0]["CodClie"].ToString();

					if(cedula == "7420758" || cedula == "12698405" || cedula == "4373147")
						cedula = "xxx";

					if(Analizar_TipoCxc(str,ref Cuota,ref Total))
					{
						Resumen = new Factura(str = idFactura);
						Resumen.Cliente = dtCxC.Rows[0]["Descrip"].ToString();

						for(int i=0;i<dtCxC.Rows.Count;i++)
						{
							str = dtCxC.Rows[i]["NumeroD"].ToString();
							Analizar_TipoCxc(str,ref Cuota,ref Total);

							DateTime FE = Convert.ToDateTime(dtCxC.Rows[i]["FechaE"].ToString());

							Dias = Analizar_FechaV(dtCxC.Rows[i]["FechaV"].ToString(),_LaFecha);
							strMonto = dtCxC.Rows[i]["Saldo"].ToString();
							Monto = Convert.ToDouble(strMonto);

							if(Monto!=0.0 || Dias == -1)
							{
								if (Dias == -1)
									Resumen.Experiencia = 0;
								else if (Dias>=0 && Dias<30)
									Resumen.Experiencia = 2;
								else if (Dias>=30 && Dias<60)
									Resumen.Experiencia = 3;
								else if (Dias>=60 && Dias<90)
									Resumen.Experiencia = 4;
								else if (Dias>=90 && Dias<120)
									Resumen.Experiencia = 5;
								else if (Dias>=120)
									Resumen.Experiencia = 20;
								else if (Dias < 0 && FE.Month != _LaFecha.Month && FE.Year != _LaFecha.Year)
									Resumen.Experiencia = 1;
								break;
							}
							else 
							{
								Resumen.Experiencia = 1;


								if(Cuota==Total)
								{
									DataTable dtFechaC = new DataTable();
									string F = "";

									DateTime FV = Convert.ToDateTime(dtCxC.Rows[i]["FechaV"].ToString());

									SQL= "SELECT SAACXC.FechaV " +
										" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
										" WHERE SAACXC.CodClie='" + idCliente + "' AND SAACXC.TipoCxc='41' AND DATEDIFF(DD,SAACXC.FechaE, CAST('" + FV.ToString("yyMMdd") + "' AS DATETIME)) <= 0  ORDER BY SAACXC.FechaV";
									clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtFechaC);
									if(dtFechaC.Rows.Count>0)
									{
										try
										{
											F = dtFechaC.Rows[0][0].ToString();
											if (F.Trim()!="") Resumen.FechaCancelacion = Convert.ToDateTime(F);
										}
										catch
										{Mensajes.Mensaje.Error(dtFechaC.Rows[0][0].ToString(),"Saint");}
										
									}
								}
								//else  Resumen.FechaCancelacion = DateTime.Now;
							}
							
						}
						return Resumen;
					}
					else return Resumen;
				}
			}
			else return Resumen;
		}


		public Documentos[] DocumentosSaint(System.Windows.Forms.ListBox list)
		{
			

			DataTable dtFact = new DataTable();
			DataTable dtCxC = new DataTable();
			DataTable dtDetFact = new DataTable();

			list.Items.Add("Obteniendo Datos Saint.....");

			/*
			 * 1.- Buscar el listado de cedulas que tienen CxC (SAACXC)
			 * 2.- Identificar por cada Cedula si es un cliente que viene del Saint Viejo (DOS)
			 * 2.1.- Si viene del saint viejo crear una factura ficticia para cada transaccion
			 * 2.2.- Si viene del Saint Nuevo (Windows) buscar la información de la factura (SAFACT) y
			 *			el detalle de la factura (SAITEMFAC)
			 * 3.- Cargar las cuotas de cada factura del cliente (SAACXC)
			 * 
			 * */

			#region SQL Union Facturas

			SQL =	" SELECT " +
				" FacturasViejas.Cedula AS CodClie, " +
				" SACLIE.Descrip, " +
				" SACLIE.Direc1, " +
				" SACLIE.Direc2, " +
				" SACLIE.Telef, " +
				" cast(FacturasViejas.NFactura as varchar(10)) AS NumeroD, " +
				" FacturasViejas.GNumero AS NGiros, " +
				" FacturasViejas.FechaE, " +
				" FacturasViejas.FechaV, " +
				" (FacturasViejas.MontoV * 0.14) AS MtoTax, " +
				" CAST(FacturasViejas.MontoV AS decimal(13)) AS MtoFinanc, " +	//Revisar
				" (CAST(FacturasViejas.MontoV AS decimal(13)) - (CAST(FacturasViejas.MontoV AS decimal(13)) * 0.14) - (CAST(FacturasViejas.MontoV AS decimal(13)) * 0.47)) AS Monto, " +						//Revisar
				" 0 as CancelE, " +
				" (CAST(FacturasViejas.MontoV AS decimal(13)) * 0.47) AS  MtoInt1" +
				" FROM " +
				" FacturasViejas INNER JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) " +
				" ON (FacturasViejas.Cedula=SACLIE.CodClie) AND ((FacturasViejas.FechaE) = (SAACXC.FechaE))  " +
				" WHERE " +
				" FacturasViejas.GNumero<>'0'  AND  " +
				" FacturasViejas.MontoV>0 " +
				//" FacturasViejas.MontoV>0  AND (FacturasViejas.Cedula = '10686579' OR FacturasViejas.Cedula = '10636307' OR FacturasViejas.Cedula = '10627872' OR FacturasViejas.Cedula = '10566379' ) " +
				" GROUP BY " +
				" FacturasViejas.Cedula, " +
				" FacturasViejas.NFactura, " +
				" FacturasViejas.GNumero, " +
				" FacturasViejas.FechaE, " +
				" FacturasViejas.FechaV, " +
				" FacturasViejas.MontoV, " +
				" SACLIE.Descrip, " +
				" SACLIE.Direc1, " +
				" SACLIE.Direc2, " +
				" SACLIE.Telef " +
				" UNION " +
				" SELECT " +
				" SAFACT.CodClie, " +
				" SACLIE.Descrip, " +
				" SACLIE.Direc1, " +
				" SACLIE.Direc2, " +
				" SACLIE.Telef, " +
				" SAFACT.NumeroD, " +
				" SAFACT.NGiros, " +
				" SAFACT.FechaE, " +
				" SAFACT.FechaV, " +
				" SAFACT.MtoTax, " +
				" SAFACT.MtoFinanc,  " +
				" SAFACT.Monto, " +
				" SAFACT.CancelE, " +
				" SAFACT.MtoInt1 " +
				" FROM " +
				" SAFACT INNER JOIN " +
				" (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) ON " +
				" ((SAFACT.CodClie=SACLIE.CodClie) AND DATEDIFF(dd,SAFACT.FechaE,SAACXC.FechaE)=0) " +
				" WHERE " +
				" SAFACT.NGiros>0  AND " +
				//" SAFACT.MtoFinanc >0 AND (SAFACT.CodClie = '10686579' OR SAFACT.CodClie = '10636307' OR SAFACT.CodClie = '10627872' OR SAFACT.CodClie = '10566379' ) " +
				" SAFACT.MtoFinanc >0 " +
				" GROUP BY " +
				" SAFACT.CodClie, " +
				" SAFACT.NumeroD, " +
				" SAFACT.NGiros, " +
				" SAFACT.FechaE, " +
				" SAFACT.FechaV, " +
				" SAFACT.MtoTax, " +
				" SAFACT.MtoFinanc, " +
				" SAFACT.Monto, " +
				" SAFACT.CancelE, " +
				" SAFACT.MtoInt1, " +
				" SACLIE.Descrip, " +
				" SACLIE.Direc1, " +
				" SACLIE.Direc2, " +
				" SACLIE.Telef  " +
				" ORDER BY " +
				" SAFACT.NumeroD; ";

			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtFact);

			#endregion

			Documentos[] saints = new Documentos[dtFact.Rows.Count];

			list.Items.Add("Obteniendo " + dtFact.Rows.Count + " Registros de Saint." );
			list.Refresh();

			for(int i=0;i<dtFact.Rows.Count;i++)
			{
				#region SQL CxC
				SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
					" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
					" WHERE SAACXC.TipoCxc='60' AND DATEDIFF(dd,'" + Convert.ToDateTime(dtFact.Rows[i]["FechaE"].ToString()).ToString("yyMMdd") + "',SAACXC.FechaE)=0 AND SAACXC.CodClie = '" + dtFact.Rows[i]["CodClie"].ToString() + "' " +
					" ORDER BY SAACXC.NroUnico ASC; ";
				#endregion
				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

				saints[i].fact.IDCliente = dtFact.Rows[i]["CodClie"].ToString();
				saints[i].fact.Cliente = dtFact.Rows[i]["Descrip"].ToString();
				saints[i].fact.Direccion = dtFact.Rows[i]["Direc1"].ToString() + " " + dtFact.Rows[i]["Direc2"].ToString();
				saints[i].fact.Telefono = dtFact.Rows[i]["Telef"].ToString();

				saints[i].fact.NroFactura = dtFact.Rows[i]["NumeroD"].ToString();
				saints[i].fact.Giros = Convert.ToInt32(dtFact.Rows[i]["NGiros"].ToString());

				saints[i].fact.MontoTotal = Convert.ToDouble(dtFact.Rows[i]["MtoFinanc"].ToString());
				saints[i].fact.Impuesto = Convert.ToDouble(dtFact.Rows[i]["MtoTax"].ToString());
				saints[i].fact.MontoNeto = Convert.ToDouble(dtFact.Rows[i]["Monto"].ToString());
				saints[i].fact.Intereses = Convert.ToDouble(dtFact.Rows[i]["MtoInt1"].ToString());
				saints[i].fact.Adelanto = Convert.ToDouble(dtFact.Rows[i]["CancelE"].ToString());
				saints[i].fact.PagoMensual = saints[i].fact.MontoTotal /saints[i].fact.Giros;

				saints[i].fact.FechaE = Convert.ToDateTime(dtFact.Rows[i]["FechaE"].ToString());
				saints[i].fact.FechaCancelacion = Convert.ToDateTime(dtFact.Rows[i]["FechaV"].ToString());

				saints[i].cxc = new CxC[dtCxC.Rows.Count];

				int total=0;
				for(int k=0;k<dtCxC.Rows.Count;k++)
				{
					saints[i].cxc[k].FechaE = Convert.ToDateTime(dtCxC.Rows[k]["FechaE"].ToString());
					saints[i].cxc[k].FechaV = Convert.ToDateTime(dtCxC.Rows[k]["FechaV"].ToString());

					saints[i].cxc[k].Monto = Convert.ToDouble(dtCxC.Rows[k]["Monto"].ToString());

					Analizar_TipoCxc(dtCxC.Rows[k]["NumeroD"].ToString(),ref saints[i].cxc[k].NroCuota,ref total);

					saints[i].cxc[k].Saldo = Convert.ToDouble(dtCxC.Rows[k]["Saldo"].ToString());

					if(saints[i].cxc[k].Saldo == 0.0) saints[i].cxc[k].Cancelada = true;
					else saints[i].cxc[k].Cancelada = false;

				}

				#region SQL Items
				SQL= "SELECT SAITEMFAC.CodItem,SAITEMFAC.Descrip1,SAITEMFAC.Precio,SAITEMFAC.Costo " +
					" FROM SAITEMFAC " +
					" WHERE SAITEMFAC.NumeroD = '" + saints[i].fact.NroFactura + "'";
				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtDetFact);
				#endregion

				if(dtDetFact.Rows.Count>0)
				{
					saints[i].detfact = new DetalleFactura[dtDetFact.Rows.Count];
					for(int j=0;j<dtDetFact.Rows.Count;j++)
					{
						saints[i].detfact[j].CodItem = dtDetFact.Rows[j]["CodItem"].ToString();
						saints[i].detfact[j].Costo = Convert.ToDouble(dtDetFact.Rows[j]["Costo"].ToString());
						saints[i].detfact[j].Descripcion = dtDetFact.Rows[j]["Descrip1"].ToString();
					}

				}
				else
				{
					saints[i].detfact = new DetalleFactura[1];
					saints[i].detfact[0].CodItem = "ARTGEN01";
					saints[i].detfact[0].Costo = 1.0;
					saints[i].detfact[0].Descripcion = "Artículo Genérico";

				}

				//list.Items.Add("Factura => " + saints[i].fact.NroFactura + " Cliente => " + saints[i].fact.IDCliente);
				//list.Refresh();
				//list.SendToBack();
			}
			list.Items.Add("Registros de Saint Obtenidos." );
			return saints;
		}


		public static void ExportarLogs(string[] lineas,string arch)
		{
			if (System.IO.File.Exists(arch)){System.IO.File.Delete(arch);}
			StreamWriter TxtFile = new StreamWriter (arch,true);

			for(int i=0;i<lineas.Length;i++)
			{
				string Cad = lineas[i];

				TxtFile.WriteLine(Cad);
			}

			TxtFile.Close();

		}


		private ResumenCxC CxCCliente(string idCliente)
		{
			DataTable dtCxC = new DataTable();
			string str="";
			ResumenCxC Resumen;

			SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
				" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
				" WHERE SAACXC.CodClie='" + idCliente + "' AND SAACXC.TipoCxc='60' AND SAACXC.TipoCxc='10' " +
				" ORDER BY SAACXC.NroUnico ASC; ";
			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

			if(dtCxC.Rows.Count>0)
			{
				if(dtCxC.Rows.Count>20)
				{
					str = dtCxC.Rows[0][0].ToString();

					return new ResumenCxC("");
				}
				else
				{
					int Cuota=0;
					int Total=0;
					int Dias =0;
					double Monto=0.0;
					string strMonto;
					str = dtCxC.Rows[0]["NumeroD"].ToString();
					if(Analizar_TipoCxc(str,ref Cuota,ref Total))
					{
						Resumen = new ResumenCxC(str = dtCxC.Rows[0]["CodClie"].ToString());
						Resumen.Cliente = dtCxC.Rows[0]["Descrip"].ToString();
						Resumen.FechaC = dtCxC.Rows[0]["FechaE"].ToString().Substring(0,10);
						Resumen.Cuotas = Total;

						for(int i=0;i<dtCxC.Rows.Count;i++)
						{
							Dias = Analizar_FechaV(dtCxC.Rows[i]["FechaV"].ToString(),_LaFecha);
							strMonto = dtCxC.Rows[i]["Saldo"].ToString();
							Monto = Convert.ToDouble(strMonto);

							if(Dias==-1 && Monto!=0.0) 
							{
								Resumen.CuotasVencidas++;
								Resumen.Vencido += Monto;
							}
							else if(Dias!=-1 && Monto!=0.0)
							{
								Resumen.CuotasxVencer++;
								if (Dias>=0 && Dias<30)
									Resumen.xVencer0_30 += Monto;
								else if (Dias>=30 && Dias<60)
									Resumen.xVencer30_60 += Monto;
								else if (Dias>=60 && Dias<90)
									Resumen.xVencer60_90 += Monto;
								else if (Dias>=90 && Dias<120)
									Resumen.xVencer90_120 += Monto;
								else if(Dias>=120)
									Resumen.xVencerMayor120 += Monto;
							}
						}

						Resumen.Total = Resumen.Vencido + Resumen.xVencer0_30 + Resumen.xVencer30_60 + Resumen.xVencer60_90 + Resumen.xVencer90_120 + Resumen.xVencerMayor120;

						return Resumen;
					}
					else return new ResumenCxC("");
				}
			}
			else return new ResumenCxC("");
		}


		private bool Analizar_TipoCxc(string Tipo,ref int Cuota,ref int Total)
		{
			Tipo = Tipo.Trim();

			int Tam = Tipo.Length;
			bool Barra =false;
			bool De = false;

			int PosBarra=0;
			int PosDe=0;
			string str;

			PosBarra = Tipo.IndexOf(Convert.ToChar(@"/"));
			if(PosBarra !=-1) Barra =true;

			PosDe = Tipo.IndexOf(Convert.ToChar("d"));
			if(PosDe!=-1) De =true;

			if((Tam==5 || Tam==7) && Barra)
			{
				str = Tipo.Substring(0,PosBarra);
				Cuota = Convert.ToInt32(str);

				str = Tipo.Substring(PosBarra+1);
				Total = Convert.ToInt32(str);
				return true;
			}
			else if((Tam==6 || Tam==8) && De)
			{
				str = Tipo.Substring(0,PosDe);
				Cuota = Convert.ToInt32(str);

				str = Tipo.Substring(PosDe+2);
				Total = Convert.ToInt32(str);
				return true;
			}
			else
				return false;
		}

		private int Analizar_FechaV(string FechaV,DateTime FechaActual)
		{
			DateTime Fecha = Convert.ToDateTime(FechaV);

			TimeSpan Tiempo = FechaActual.Subtract(Fecha);

			if((int)Tiempo.TotalDays < 0) return -1;
			else return (int)Tiempo.TotalDays;
		}


		#endregion

	}
}
