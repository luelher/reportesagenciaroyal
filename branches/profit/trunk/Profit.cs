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
using GrupoEmporium.Reportes;

namespace GrupoEmporium.Profit.Reportes
{
	/// <summary>
	/// Descripción breve de Profit.
	/// </summary>
	public class Profit
	{

		#region Variables

		string SQL="";
        string SQL2 = "";
		DateTime _LaFecha;

		GrupoEmporium.Varias.ClaseDocumentosXML configxml;

		clsBDConexion strConexion;
		string cadenaconexion;
		clsBDConexion strConexion_Profit;
		string cadenaconexion_Profit;
		clsBDConexion strConexion_Profit_1;
		string cadenaconexion_Profit_1;
		
		//Este objeto conexion debe ser reemplazado al cambiar a SQL Server
		public OleDbConnection Conexion;
		public OleDbConnection Conexion_Profit;
		public OleDbConnection Conexion_Profit_1;


		private static string Tab="	";
		private static char enter = (char)13;
		private static char retorno = (char)10;
        private static char espacio = (char)32;

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
			public int Dias;
			public bool Cancelada;
			public double Saldo;
			public double SaldoVencido;
			public double SaldoRestante;
			public int GirosSinCancelar;
			public double SaldoVencidoSinCancelar;
			public int GirosVencidosSinCancelar;

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
				Dias = 0;
				GirosSinCancelar = 0;
				Cancelada = false;
				Saldo = 0.0;
				SaldoVencido = 0.0;
				SaldoRestante = 0.0;
				SaldoVencidoSinCancelar = 0.0;
				GirosVencidosSinCancelar = 0;

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
		public Profit()
		{
			CargarConfig(true);

			_LaFecha = DateTime.Now;

			try
			{
				Conexion.Open();
				//Mensajes.Mensaje.Error(Conexion.State.ToString(),"Profit Reportes");
			}
			catch(Exception ex)
			{
                Mensajes.Mensaje.Error(ex.Message, "Profit Reportes");
			}

		}
		public Profit(DateTime Fecha)
		{

			CargarConfig(true);

			_LaFecha = Fecha;

			try
			{
				Conexion.Open();
                //Mensajes.Mensaje.Error(Conexion.State.ToString(), "Profit Reportes");
			}
			catch(Exception ex)
			{
                Mensajes.Mensaje.Error(ex.Message, "Profit Reportes");
			}

		}


		public Profit(DateTime Fecha, bool conex)
		{

			CargarConfig(conex);

			_LaFecha = Fecha;

			try
			{
				Conexion.Open();
                //Mensajes.Mensaje.Error(Conexion.State.ToString(), "Profit Reportes");
			}
			catch(Exception ex)
			{
                Mensajes.Mensaje.Error(ex.Message, "Profit Reportes");
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

		public DataTable Reporte_Morosos_Todos()
		{
			int desde = 1; int hasta = 40;
			//Conexion_Profit_1.Open();

			if(Conexion.State == ConnectionState.Open)
			{

				DataTable dt = new DataTable();
				DataTable dtReporte = new DataTable();
				DataTable dt_nro_doc = new DataTable();
				DataTable dt_fechas_cobros = new DataTable();
				DataRow dr;
				string cliente = "";
				bool hecho = false;
				int dias_ultimo_pago=0;
				Decimal meses=0;
				int resto = 0;

				DateTime fecha_ultimo_cobro;
			
				#region SQL Union Facturas

				SQL =	" SELECT " +
					" docum_cc.co_cli     AS CodClie, " +
					" clientes.cli_des    AS Descrip, " +
					" clientes.direc1     AS Direc1, " +
					" clientes.direc2     AS Direc2, " +
					" clientes.telefonos  AS Telef, " +
					" docum_cc.nro_doc    AS NumeroD, " +
					" condicio.dias_cred  AS NGiros, " +
					" docum_cc.fec_emis   AS FechaE, " +
					" docum_cc.fec_venc   AS FechaV, " +
					" docum_cc.monto_net  AS MtoFinanc " +
					" FROM " +
					" ((docum_cc INNER JOIN clientes ON docum_cc.co_cli = clientes.co_cli) " +
					" INNER JOIN (factura INNER JOIN condicio ON factura.forma_pag = condicio.co_cond) ON docum_cc.nro_doc = factura.fact_num) " +
					" WHERE " +
					" docum_cc.tipo_doc = 'FACT' " +
                    //" AND condicio.dias_cred > 0 " +
					//" AND docum_cc.co_cli='7983524' " + 
					" ORDER BY " +
					" CodClie ASC;";

				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

				Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

				#endregion

				if(dt.Rows.Count>0)
				{
					Factura RFactura;

					dtReporte.Columns.Add("cliente",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("cedula",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("direccion",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("telefono",System.Type.GetType("System.String"));

					dtReporte.Columns.Add("fechae",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("fechav",System.Type.GetType("System.DateTime"));

					dtReporte.Columns.Add("meses",System.Type.GetType("System.Byte"));
					dtReporte.Columns.Add("saldo",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("ultcobro",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("dias",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("pagomensual",System.Type.GetType("System.Double"));

					//dtReporte.Columns.Add("saldovencido",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("saldovencidosincancelar",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("saldorestante",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("girossincancelar",System.Type.GetType("System.Byte"));
					dtReporte.Columns.Add("girosvencidossincancelar",System.Type.GetType("System.Byte"));
					//dtReporte.Columns.Add("impuesto",System.Type.GetType("System.Double"));
					//dtReporte.Columns.Add("intereses",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("diasultimopago",System.Type.GetType("System.String"));

					for(int i=0;i<dt.Rows.Count;i++)
					{

						SQL =	"SELECT " +
							" docum_cc.nro_doc " +
							"FROM " +
							" docum_cc " +
							"WHERE " +
							" docum_cc.tipo_doc = 'CFXG' AND docum_cc.fec_emis = '" + Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString()).ToString("yyyy-MM-dd") + "' AND docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "'";
						clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_nro_doc);

						if(cliente != dt.Rows[i]["CodClie"].ToString())
						{
							cliente = dt.Rows[i]["CodClie"].ToString();
							hecho = true;
						}
						
						if(dt_nro_doc.Rows.Count>0 && hecho)
						{
							RFactura = ResumenFactura(dt_nro_doc.Rows[0]["nro_doc"].ToString(),dt.Rows[i]["CodClie"].ToString(),(DateTime)dt.Rows[i]["FechaE"]);

							if(!RFactura.Cancelada)
							{
								RFactura.IDCliente = dt.Rows[i]["CodClie"].ToString();
								int CantGiros = Convert.ToInt32( dt.Rows[i]["NGiros"].ToString() )/30;
								//if(CantGiros>0) RFactura.PagoMensual  = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString())/CantGiros;

								RFactura.Giros = CantGiros;
								//RFactura.MontoTotal = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString());
								RFactura.FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
								//RFactura.FechaCancelacion = Convert.ToDateTime(dt.Rows[i]["FechaV"].ToString());
								RFactura.Cliente = dt.Rows[i]["Descrip"].ToString();

								SQL =	" SELECT cobros.fec_cob  " +
									" FROM " +
									" docum_cc INNER JOIN (reng_cob INNER JOIN cobros ON reng_cob.cob_num = cobros.cob_num) ON docum_cc.nro_doc = reng_cob.doc_num " +
									" WHERE " +
									" docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "' " +
									" ORDER BY " +
									" cobros.fec_cob DESC ";
								clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_fechas_cobros);

								if(dt_fechas_cobros.Rows.Count>0)
								{
									fecha_ultimo_cobro=Convert.ToDateTime(dt_fechas_cobros.Rows[0][0].ToString());}
								else{fecha_ultimo_cobro=DateTime.MinValue;}

								dias_ultimo_pago = Analizar_FechaV(fecha_ultimo_cobro.ToString(),DateTime.Now);;

								resto = RFactura.Dias % 30;
								meses = Math.Abs(RFactura.Dias / 30);
								if(resto>0) meses++;

								if(RFactura.NroFactura !="" && (RFactura.GirosVencidosSinCancelar >=desde && RFactura.GirosVencidosSinCancelar <=hasta))
								{
									dr = dtReporte.NewRow();
									dr["cliente"] = RFactura.Cliente;
									dr["cedula"] = RFactura.IDCliente;
									dr["direccion"] = dt.Rows[i]["Direc1"].ToString();
									dr["telefono"] = dt.Rows[i]["Telef"].ToString();
									dr["fechae"] = dt.Rows[i]["FechaE"].ToString();
									dr["fechav"] = dt.Rows[i]["FechaV"].ToString();

									dr["meses"] = RFactura.GirosVencidosSinCancelar;

									dr["saldo"] = RFactura.SaldoVencidoSinCancelar;
									dr["ultcobro"] = fecha_ultimo_cobro;
									dr["dias"] = dias_ultimo_pago;
									dr["pagomensual"] = RFactura.PagoMensual.ToString("#########.00");

									//dr["saldovencido"] = RFactura.SaldoVencido;
									dr["saldovencidosincancelar"] = RFactura.SaldoVencidoSinCancelar;
									dr["saldorestante"] = RFactura.SaldoRestante;
									dr["girossincancelar"] = RFactura.GirosSinCancelar;
									dr["girosvencidossincancelar"] = RFactura.GirosVencidosSinCancelar;
									//dr["impuesto"] = RFactura.Impuesto;
									//dr["intereses"] = RFactura.Intereses;
									dr["diasultimopago"] = dias_ultimo_pago;

									dtReporte.Rows.Add(dr);
									hecho = false;
								}
							}
						}
					}
					dtReporte.AcceptChanges();
					CerrarConexiones();
					return dtReporte;
				}
				else {CerrarConexiones(); return new DataTable();}
			}
			else {CerrarConexiones(); return new DataTable();}
		}

		public DataTable Reporte_Morosos(int desde, int hasta )
		{
			//Conexion.Open();

			if(Conexion.State == ConnectionState.Open)
			{

				DataTable dt = new DataTable();
				DataTable dtReporte = new DataTable();
				DataTable dt_nro_doc = new DataTable();
				DataTable dt_fechas_cobros = new DataTable();
				DataRow dr;
				string cliente = "";
				bool hecho = false;
				int dias_ultimo_pago=0;
				Decimal meses=0;
				int resto = 0;

				DateTime fecha_ultimo_cobro;
			
				#region SQL Union Facturas

				SQL =	" SELECT " +
					" docum_cc.co_cli     AS CodClie, " +
					" clientes.cli_des    AS Descrip, " +
					" clientes.direc1     AS Direc1, " +
					" clientes.direc2     AS Direc2, " +
					" clientes.telefonos  AS Telef, " +
					" docum_cc.nro_doc    AS NumeroD, " +
					" condicio.dias_cred  AS NGiros, " +
					" docum_cc.fec_emis   AS FechaE, " +
					" docum_cc.fec_venc   AS FechaV, " +
					" docum_cc.monto_net  AS MtoFinanc " +
					" FROM " +
					" ((docum_cc INNER JOIN clientes ON docum_cc.co_cli = clientes.co_cli) " +
					" INNER JOIN (factura INNER JOIN condicio ON factura.forma_pag = condicio.co_cond) ON docum_cc.nro_doc = factura.fact_num) " +
					" WHERE " +
					" docum_cc.tipo_doc = 'FACT' " +
                    //" AND condicio.dias_cred > 0 " +
					//" AND docum_cc.co_cli='7983524' " + 
					" ORDER BY " +
					" CodClie ASC;";

				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

				Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

				#endregion

				if(dt.Rows.Count>0)
				{
					Factura RFactura;

					dtReporte.Columns.Add("cliente",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("cedula",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("direccion",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("telefono",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("meses",System.Type.GetType("System.Byte"));
					dtReporte.Columns.Add("saldo",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("ultcobro",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("dias",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("pagomensual",System.Type.GetType("System.String"));

					for(int i=0;i<dt.Rows.Count;i++)
					{

						SQL =	"SELECT " +
							" docum_cc.nro_doc " +
							"FROM " +
							" docum_cc " +
							"WHERE " +
							" docum_cc.tipo_doc = 'CFXG' AND docum_cc.fec_emis = '" + Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString()).ToString("yyyy-MM-dd") + "' AND docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "'";
						clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_nro_doc);

						if(cliente != dt.Rows[i]["CodClie"].ToString())
						{
							cliente = dt.Rows[i]["CodClie"].ToString();
							hecho = true;
						}
						
						if(dt_nro_doc.Rows.Count>0 && hecho)
						{
							RFactura = ResumenFactura(dt_nro_doc.Rows[0]["nro_doc"].ToString(),dt.Rows[i]["CodClie"].ToString(),(DateTime)dt.Rows[i]["FechaE"]);

							if(!RFactura.Cancelada)
							{
								RFactura.IDCliente = dt.Rows[i]["CodClie"].ToString();
								int CantGiros = Convert.ToInt32( dt.Rows[i]["NGiros"].ToString() )/30;
								//if(CantGiros>0) RFactura.PagoMensual  = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString())/CantGiros;

								//RFactura.Giros = CantGiros;
								//RFactura.MontoTotal = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString());
								RFactura.FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
								//RFactura.FechaCancelacion = Convert.ToDateTime(dt.Rows[i]["FechaV"].ToString());
								RFactura.Cliente = dt.Rows[i]["Descrip"].ToString();

								SQL =	" SELECT cobros.fec_cob  " +
									" FROM " +
									" docum_cc INNER JOIN (reng_cob INNER JOIN cobros ON reng_cob.cob_num = cobros.cob_num) ON docum_cc.nro_doc = reng_cob.doc_num " +
									" WHERE " +
									" docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "' " +
									" ORDER BY " +
									" cobros.fec_cob DESC ";
								clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_fechas_cobros);

								if(dt_fechas_cobros.Rows.Count>0)
								{
									fecha_ultimo_cobro=Convert.ToDateTime(dt_fechas_cobros.Rows[0][0].ToString());}
								else{fecha_ultimo_cobro=DateTime.MinValue;}

								dias_ultimo_pago = Analizar_FechaV(fecha_ultimo_cobro.ToString(),DateTime.Now);;

								resto = RFactura.Dias % 30;
								meses = Math.Abs(RFactura.Dias / 30);
								if(resto>0) meses++;

								if(RFactura.NroFactura !="" && (RFactura.GirosVencidosSinCancelar >=desde && RFactura.GirosVencidosSinCancelar <=hasta) && dias_ultimo_pago > 30 )
								{
									dr = dtReporte.NewRow();
									dr["cliente"] = RFactura.Cliente;
									dr["cedula"] = RFactura.IDCliente;
									dr["direccion"] = dt.Rows[i]["Direc1"].ToString();
									dr["telefono"] = dt.Rows[i]["Telef"].ToString();

									dr["meses"] = RFactura.GirosVencidosSinCancelar;

									dr["saldo"] = RFactura.SaldoVencidoSinCancelar;
									dr["ultcobro"] = fecha_ultimo_cobro;
									dr["dias"] = dias_ultimo_pago;
									dr["pagomensual"] = RFactura.PagoMensual.ToString("#########.00");

									dtReporte.Rows.Add(dr);
									hecho = false;
								}
							}
						}
					}
					dtReporte.AcceptChanges();
					CerrarConexiones();
					return dtReporte;
				}else {CerrarConexiones(); return new DataTable();}
			}
			else {CerrarConexiones(); return new DataTable();}
		}


		public DataTable Reporte_Gerencial_Morosos(int desde, int hasta)
		{
			//Conexion.Open();

			if(Conexion.State == ConnectionState.Open)
			{

				DataTable dt = new DataTable();
				DataTable dtReporte = new DataTable();
				DataTable dt_nro_doc = new DataTable();
				DataTable dt_fechas_cobros = new DataTable();
				DataRow dr;
				string cliente = "";
				bool hecho = false;
				int dias_ultimo_pago=0;
				Decimal meses=0;
				int resto = 0;

				DateTime fecha_ultimo_cobro;
			
				#region SQL Union Facturas

				SQL =	" SELECT " +
					" docum_cc.co_cli     AS CodClie, " +
					" clientes.cli_des    AS Descrip, " +
					" clientes.direc1     AS Direc1, " +
					" clientes.direc2     AS Direc2, " +
					" clientes.telefonos  AS Telef, " +
					" docum_cc.nro_doc    AS NumeroD, " +
					" condicio.dias_cred  AS NGiros, " +
					" docum_cc.fec_emis   AS FechaE, " +
					" docum_cc.fec_venc   AS FechaV, " +
					" docum_cc.monto_net  AS MtoFinanc " +
					" FROM " +
					" ((docum_cc INNER JOIN clientes ON docum_cc.co_cli = clientes.co_cli) " +
					" INNER JOIN (factura INNER JOIN condicio ON factura.forma_pag = condicio.co_cond) ON docum_cc.nro_doc = factura.fact_num) " +
					" WHERE " +
					" docum_cc.tipo_doc = 'FACT' " +
                    //" AND condicio.dias_cred > 0 " +
					//" AND docum_cc.co_cli='7308142' " + 
					" ORDER BY " +
					" CodClie ASC;";

				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

				Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

				#endregion

				if(dt.Rows.Count>0)
				{
					Factura RFactura;

					dtReporte.Columns.Add("cliente",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("cedula",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("direccion",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("telefono",System.Type.GetType("System.String"));
					dtReporte.Columns.Add("meses",System.Type.GetType("System.Byte"));
					dtReporte.Columns.Add("saldo",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("ultcobro",System.Type.GetType("System.DateTime"));
					dtReporte.Columns.Add("dias",System.Type.GetType("System.Double"));
					dtReporte.Columns.Add("pagomensual",System.Type.GetType("System.String"));

					for(int i=0;i<dt.Rows.Count;i++)
					{

						SQL =	"SELECT " +
							" docum_cc.nro_doc " +
							"FROM " +
							" docum_cc " +
							"WHERE " +
							" docum_cc.tipo_doc = 'CFXG' AND docum_cc.fec_emis = '" + Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString()).ToString("yyyy-MM-dd") + "' AND docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "'";
						clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_nro_doc);

						if(cliente != dt.Rows[i]["CodClie"].ToString())
						{
							cliente = dt.Rows[i]["CodClie"].ToString();
							hecho = true;
						}
						

						if(dt_nro_doc.Rows.Count>0 && hecho)
						{
							RFactura = ResumenFactura(dt_nro_doc.Rows[0]["nro_doc"].ToString(),dt.Rows[i]["CodClie"].ToString(),(DateTime)dt.Rows[i]["FechaE"]);

							if(!RFactura.Cancelada)
							{
								RFactura.IDCliente = dt.Rows[i]["CodClie"].ToString();
								int CantGiros = Convert.ToInt32( dt.Rows[i]["NGiros"].ToString() )/30;
								//if(CantGiros>0) RFactura.PagoMensual  = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString())/CantGiros;

								//RFactura.Giros = CantGiros;
								//RFactura.MontoTotal = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString());
								RFactura.FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
								//RFactura.FechaCancelacion = Convert.ToDateTime(dt.Rows[i]["FechaV"].ToString());
								RFactura.Cliente = dt.Rows[i]["Descrip"].ToString();

								SQL =	" SELECT cobros.fec_cob  " +
									" FROM " +
									" docum_cc INNER JOIN (reng_cob INNER JOIN cobros ON reng_cob.cob_num = cobros.cob_num) ON docum_cc.nro_doc = reng_cob.doc_num " +
									" WHERE " +
									" docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "' " +
									" ORDER BY " +
									" cobros.fec_cob DESC ";
								clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt_fechas_cobros);

								if(dt_fechas_cobros.Rows.Count>0)
								{
									fecha_ultimo_cobro=Convert.ToDateTime(dt_fechas_cobros.Rows[0][0].ToString());}
								else{fecha_ultimo_cobro=DateTime.MinValue;}

								dias_ultimo_pago = Analizar_FechaV(fecha_ultimo_cobro.ToString(),DateTime.Now);;

								resto = RFactura.Dias % 30;
								meses = Math.Abs(RFactura.Dias / 30);
								if(resto>0) meses++;

								if(RFactura.NroFactura !="" && (RFactura.GirosSinCancelar >=desde && RFactura.GirosSinCancelar <=hasta) && dias_ultimo_pago > 30 )
								{
									dr = dtReporte.NewRow();
									dr["cliente"] = RFactura.Cliente;
									dr["cedula"] = RFactura.IDCliente;
									dr["direccion"] = dt.Rows[i]["Direc1"].ToString();
									dr["telefono"] = dt.Rows[i]["Telef"].ToString();

									dr["meses"] = RFactura.GirosVencidosSinCancelar;

									dr["saldo"] = RFactura.SaldoVencidoSinCancelar;
									dr["ultcobro"] = fecha_ultimo_cobro;
									dr["dias"] = dias_ultimo_pago;
									dr["pagomensual"] = RFactura.PagoMensual.ToString("#########.00");

									dtReporte.Rows.Add(dr);
									hecho = false;
								}
							}
						}
					}
					dtReporte.AcceptChanges();
					CerrarConexiones();
					return dtReporte;
				}
				else {CerrarConexiones(); return new DataTable();}
			}
			else {CerrarConexiones(); return new DataTable();}
		}


		public DataTable Reporte_Experiencia(FormReportes frm)
		{
            string str = "";
			if(Conexion.State == ConnectionState.Open)
			{
				// 641.555   10842309

				DataTable dt = new DataTable();
                DataTable dt2 = new DataTable();
				DataTable dtReporte = new DataTable();
				DataTable dt_nro_doc = new DataTable();
				DataRow dr;
			
				#region SQL Union Facturas

                SQL = " SELECT " +
                            " docum_cc.co_cli     AS CodClie, " +
                            " clientes.cli_des    AS Descrip, " +
                            " clientes.direc1     AS Direc1, " +
                            " clientes.direc2     AS Direc2, " +
                            " clientes.telefonos  AS Telef, " +
                            " docum_cc.nro_doc    AS NumeroD, " +
                            " 0  AS NGiros, " +
                            " docum_cc.fec_emis   AS FechaE, " +
                            " docum_cc.fec_venc   AS FechaV, " +
                            " docum_cc.monto_net  AS MtoFinanc " +
                        " FROM " +
                            " ((docum_cc INNER JOIN factura ON docum_cc.nro_doc = factura.fact_num) INNER JOIN clientes ON docum_cc.co_cli = clientes.co_cli) " +
                        " WHERE " +
                            //" docum_cc.tipo_doc = 'XXXX' AND " +
                            //" docum_cc.co_cli = '10963299' AND " +
                            " clientes.co_seg != 'EMPRE' AND " +
                            " (docum_cc.tipo_doc = 'FACT') AND docum_cc.co_sucu<>'ADMINI' " +
                            " AND docum_cc.fec_emis < '" + _LaFecha.ToString("yyyy-MM-dd") + "'" +
                            " AND factura.anulada = 0";

                SQL2 = " SELECT " +
	                        " docum_cc.co_cli     AS CodClie,  " +
	                        " clientes.cli_des    AS Descrip,  " +
	                        " ''     AS Direc1,  " +
	                        " ''     AS Direc2,  " +
	                        " ''  AS Telef,  " +
	                        " 0    AS NumeroD,  " +
	                        " 0  AS NGiros,  " +
	                        " docum_cc.fec_emis   AS FechaE,  " +
	                        " '2009-01-01'   AS FechaV,  " +
	                        " 0   AS MtoFinanc  " +
                        " FROM  " +
	                        " (docum_cc INNER JOIN clientes ON docum_cc.co_cli = clientes.co_cli)   " +
						" WHERE " +
                            //" docum_cc.co_cli = '10963299' AND " +
                            //" docum_cc.tipo_doc = 'XXXX' AND " +
                            " clientes.co_seg != 'EMPRE' AND " +
                            " (docum_cc.tipo_doc = 'GIRO') AND docum_cc.co_sucu<>'ADMINI' " +
                            " AND docum_cc.fec_emis < '" + _LaFecha.ToString("yyyy-MM-dd") + "' AND docum_cc.nro_orig=0 " +
                            " GROUP BY docum_cc.co_cli, clientes.cli_des, docum_cc.fec_emis ";

				clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);
                clsBD.EjecutarQuery(strConexion, Conexion, SQL2, out dt2);

                Mensajes.Mensaje.Informar((dt.Rows.Count + dt2.Rows.Count).ToString() + " Clientes a Procesar", "Saint Reportes");
                int k = 0;
                int registros = dt.Rows.Count + dt2.Rows.Count;
				#endregion

                if (dt.Rows.Count > 0 || dt2.Rows.Count > 0)
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

                    for (int j = 0; j < 2; j++)
                    {
                        if (j == 1) dt = dt2;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            k++;
                            if (frm != null)
                            {
                                frm.labelGenerando.Text = "Procesando... " + k + " de " + registros;
                                frm.Refresh();
                            } 

                            SQL = "SELECT " +
                                        " docum_cc.origen_d as nro_doc " +
                                    "FROM " +
                                        " docum_cc " +
                                    "WHERE " +
                                        " docum_cc.tipo_doc = 'CFXG' AND docum_cc.fec_emis = '" + Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString()).ToString("yyyy-MM-dd") + "' AND docum_cc.co_cli = '" + dt.Rows[i]["CodClie"].ToString() + "'";
                            clsBD.EjecutarQuery(strConexion, Conexion, SQL, out dt_nro_doc);

                            //if (dt_nro_doc.Rows.Count > 0)
                            if(true)
                            {
                                if (dt_nro_doc.Rows.Count > 0) str = dt_nro_doc.Rows[0]["nro_doc"].ToString().Trim();
                                else str = "";
                                RFactura = ResumenFactura(str, dt.Rows[i]["CodClie"].ToString().Trim(), (DateTime)dt.Rows[i]["FechaE"]);

                                RFactura.IDCliente = dt.Rows[i]["CodClie"].ToString();
                                //int CantGiros = Convert.ToInt32( dt.Rows[i]["NGiros"].ToString() )/30;
                                int CantGiros = RFactura.Giros;
                                //if(CantGiros>0) RFactura.PagoMensual  = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString())/CantGiros;

                                RFactura.Giros = CantGiros;
                                //RFactura.MontoTotal = Convert.ToDouble(dt.Rows[i]["MtoFinanc"].ToString());
                                RFactura.FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
                                //RFactura.FechaCancelacion = Convert.ToDateTime(dt.Rows[i]["FechaV"].ToString());
                                RFactura.Cliente = dt.Rows[i]["Descrip"].ToString();

                                if (RFactura.NroFactura != "")
                                {
                                    dr = dtReporte.NewRow();
                                    dr["Cedula"] = RFactura.IDCliente;
                                    dr["Nombre"] = RFactura.Cliente;
                                    dr["Telefono"] = RFactura.Telefono;
                                    dr["Factura"] = RFactura.NroFactura;
                                    dr["FechaE"] = RFactura.FechaE;
                                    try
                                    {
                                        dr["MontoTotal"] = RFactura.MontoTotal.ToString("########.##");
                                        dr["PagoMensual"] = RFactura.PagoMensual.ToString("########.##");
                                    }
                                    catch (Exception ex) { }
                                    dr["Giros"] = RFactura.Giros;
                                    dr["FechaCancelacion"] = RFactura.FechaCancelacion;
                                    dr["Experiencia"] = RFactura.Experiencia;
                                    dtReporte.Rows.Add(dr);
                                }
                            }
                        }
                    }
					dtReporte.AcceptChanges();
					CerrarConexiones();
					return dtReporte;
				}
				else {CerrarConexiones(); return new DataTable();}
			}
			else {CerrarConexiones(); return new DataTable();}
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
                    try
                    {

				        FechaE = Convert.ToDateTime(dt.Rows[i]["FechaE"].ToString());
                        if (dt.Rows[i]["MontoTotal"] is System.DBNull) MontoT = 0;
                        else MontoT = Math.Abs(Convert.ToDouble(dt.Rows[i]["MontoTotal"].ToString()));
                        if (dt.Rows[i]["PagoMensual"] is System.DBNull) Mensual = 0;
				        else Mensual= Math.Abs(Convert.ToDouble(dt.Rows[i]["PagoMensual"].ToString()));
				        FechaV = Convert.ToDateTime(dt.Rows[i]["FechaCancelacion"].ToString());

				        string Cad = dt.Rows[i]["Cedula"].ToString() + Tab +
                            dt.Rows[i]["Nombre"].ToString().Trim().Replace(enter, espacio).Replace(retorno,espacio) + Tab +
					        Tab+
					        dt.Rows[i]["Telefono"].ToString().Trim() +Tab+
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

                    }catch(Exception ex){
                        Mensajes.Mensaje.Error(ex.Message + " Cliente: " + dt.Rows[i]["Cedula"].ToString() + " - " + dt.Rows[i]["Nombre"].ToString(), "Profit Reportes");
                    }


			    }

			TxtFile.Close();

		}

		public static void ExportarMorosos(DataTable dt,string arch)
		{
			if (System.IO.File.Exists(arch)){System.IO.File.Delete(arch);}
			StreamWriter TxtFile = new StreamWriter (arch,true);

			string Cliente;
			string Cedula;
			string Direccion;
			string Telefono;
			DateTime Fechae;
			DateTime Fechav;
			string Meses;
			double Saldo;
			DateTime Ultcobro;
			string Dias;
			double Pagomensual;
			//double Saldovencido;
			double Saldovencidosincancelar;
			double Saldorestante;
			string Girossincancelar;
			string Girosvencidossincancelar;
			//double Impuesto;
			//double Intereses;
			string Diasultimopago;

			for(int i=0;i<dt.Rows.Count;i++)
			{

				Cliente					=			dt.Rows[i]["cliente"].ToString();
				Cedula					=			dt.Rows[i]["cedula"].ToString();
				Direccion				=			dt.Rows[i]["direccion"].ToString();
				Telefono				=			dt.Rows[i]["telefono"].ToString();
				Fechae					=			Convert.ToDateTime(dt.Rows[i]["fechae"].ToString());
				Fechav					=			Convert.ToDateTime(dt.Rows[i]["fechav"].ToString());

				Meses					=			dt.Rows[i]["meses"].ToString();

				Saldo					=			Convert.ToDouble(dt.Rows[i]["saldo"].ToString());
				Ultcobro				=			Convert.ToDateTime(dt.Rows[i]["ultcobro"].ToString());
				Dias					=			dt.Rows[i]["dias"].ToString();
				Pagomensual				=			Convert.ToDouble(dt.Rows[i]["pagomensual"].ToString());

				//Saldovencido			=			Convert.ToDouble(dt.Rows[i]["saldovencido"].ToString());
				Saldovencidosincancelar	=			Convert.ToDouble(dt.Rows[i]["saldovencidosincancelar"].ToString());
				Saldorestante			=			Convert.ToDouble(dt.Rows[i]["saldorestante"].ToString());
				Girossincancelar		=			dt.Rows[i]["girossincancelar"].ToString();
				Girosvencidossincancelar=			dt.Rows[i]["girosvencidossincancelar"].ToString();
				//Impuesto				=			Convert.ToDouble(dt.Rows[i]["impuesto"].ToString());
				//Intereses				=			Convert.ToDouble(dt.Rows[i]["intereses"].ToString());
				Diasultimopago			=			dt.Rows[i]["diasultimopago"].ToString();

				string Cad = Cliente + Tab + Cedula + Tab + Direccion + Tab + Telefono + Tab + Fechae.ToString("dd/MM/yyyy") + Tab + Fechav.ToString("dd/MM/yyyy") + Tab + Meses + Tab + 
					Saldo.ToString("#,##0.00;($#,##0.00);0") + Tab + Ultcobro.ToString("dd/MM/yyyy") + Tab + 
					Dias + Tab + Pagomensual.ToString("#,##0.00;($#,##0.00);0") + Tab + 
					Saldovencidosincancelar.ToString("#,##0.00;($#,##0.00);0") + Tab + Saldorestante.ToString("#,##0.00;($#,##0.00);0") + Tab + 
					Girossincancelar + Tab + Girosvencidossincancelar + Tab + Diasultimopago;
					
				TxtFile.WriteLine(Cad);

			}

			TxtFile.Close();

		}


		public static void ExportarCartas(DataTable dt,string arch)
		{
			if (System.IO.File.Exists(arch)){System.IO.File.Delete(arch);}
			StreamWriter TxtFile = new StreamWriter (arch,true);

			string Cad = "Nombre" + Tab +
					"Cedula" + Tab + 
					"Direccion" + Tab +
					"Telefono" + Tab +
					"Meses" + Tab +
					"Saldo" + Tab +
					"Fecha_Ultimo_Cobro" + Tab +
					"Dias_Ultimo_Cobro" + Tab;
			TxtFile.WriteLine(Cad);

			for(int i=0;i<dt.Rows.Count;i++)
			{

				Cad = dt.Rows[i][0].ToString() + Tab +
					dt.Rows[i][1].ToString() + Tab + 
					dt.Rows[i][2].ToString().Replace(enter.ToString(),"").Replace(retorno.ToString()," ") + Tab +
					dt.Rows[i][3].ToString() + Tab +
					dt.Rows[i][4].ToString() + Tab +
					dt.Rows[i][5].ToString() + Tab +
					Convert.ToDateTime(dt.Rows[i][6].ToString()).ToString("dd/MM/yyyy") + Tab +
					dt.Rows[i][7].ToString() + Tab;

				TxtFile.WriteLine(Cad);

			}

			TxtFile.Close();

		}


		#endregion

		#region Metodos Privados

		private void CargarConfig(bool conex)
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

			//CargarConfigProfit(conex);

		}


        //private void CargarConfigProfit(bool conex)
        //{
			
        //    ClaseDocumentosXML MiConfig = new ClaseDocumentosXML(@"configsaint.xml");
        //    if(MiConfig.Cargado)
        //    {
        //        if (MiConfig["CadenaConexion_Profit"]=="False")
        //        {
        //            if (MiConfig["Dominio_Profit"]=="True")
        //                strConexion_Profit = new clsBDConexion(MiConfig["DataSource_Profit"],MiConfig["InitialCatalog_Profit"]);
        //            else strConexion_Profit = new clsBDConexion( MiConfig["DataSource_Profit"],MiConfig["InitialCatalog_Profit"],MiConfig["Usuario_Profit"],MiConfig["Contrasena_Profit"],false);

        //            cadenaconexion_Profit = strConexion_Profit.StringConexion;
        //        }
        //        else
        //        {
        //            strConexion_Profit = new clsBDConexion();
        //            strConexion_Profit.TipoBaseDato = TipoBD.SQL_SERVER;
        //            strConexion_Profit.StringConexion = cadenaconexion_Profit;
        //            cadenaconexion_Profit = MiConfig["Conexion_Profit"];
        //        }

        //    }

        //    Conexion_Profit = new OleDbConnection(cadenaconexion_Profit);
        //    //configxml = MiConfig;
        //    CargarConfigProfit_1(conex);

        //}


        //private void CargarConfigProfit_1(bool conex)
        //{
			
        //    ClaseDocumentosXML MiConfig = new ClaseDocumentosXML(@"configsaint.xml");
        //    if(MiConfig.Cargado)
        //    {
        //        if(conex)
        //        {
        //            if (MiConfig["CadenaConexion_Profit_1"]=="False")
        //            {
        //                if (MiConfig["Dominio_Profit_1"]=="True")
        //                    strConexion_Profit_1 = new clsBDConexion(MiConfig["DataSource_Profit_1"],MiConfig["InitialCatalog_Profit_1"]);
        //                else strConexion_Profit_1 = new clsBDConexion( MiConfig["DataSource_Profit_1"],MiConfig["InitialCatalog_Profit_1"],MiConfig["Usuario_Profit_1"],MiConfig["Contrasena_Profit_1"],false);

        //                cadenaconexion_Profit_1 = strConexion_Profit_1.StringConexion;
        //            }
        //            else
        //            {
        //                strConexion_Profit_1 = new clsBDConexion();
        //                strConexion_Profit_1.TipoBaseDato = TipoBD.SQL_SERVER;
        //                strConexion_Profit_1.StringConexion = cadenaconexion_Profit_1;
        //                cadenaconexion_Profit_1 = MiConfig["Conexion_Profit_1"];
        //            }
        //        }
        //        else{
        //            if (MiConfig["CadenaConexion_Profit_2"]=="False")
        //            {
        //                if (MiConfig["Dominio_Profit_2"]=="True")
        //                    strConexion_Profit_1 = new clsBDConexion(MiConfig["DataSource_Profit_2"],MiConfig["InitialCatalog_Profit_2"]);
        //                else strConexion_Profit_1 = new clsBDConexion( MiConfig["DataSource_Profit_2"],MiConfig["InitialCatalog_Profit_2"],MiConfig["Usuario_Profit_2"],MiConfig["Contrasena_Profit_2"],false);

        //                cadenaconexion_Profit_1 = strConexion_Profit_1.StringConexion;
        //            }
        //            else
        //            {
        //                strConexion_Profit_1 = new clsBDConexion();
        //                strConexion_Profit_1.TipoBaseDato = TipoBD.SQL_SERVER;
        //                strConexion_Profit_1.StringConexion = cadenaconexion_Profit_1;
        //                cadenaconexion_Profit_1 = MiConfig["Conexion_Profit_2"];
        //            }

        //        }

        //    }

        //    Conexion_Profit_1 = new OleDbConnection(cadenaconexion_Profit_1);
        //    //configxml = MiConfig;

        //}

		private void CerrarConexiones()
		{
			if(Conexion.State == ConnectionState.Open) Conexion.Close();
			//if(Conexion_Profit.State == ConnectionState.Open) Conexion_Profit.Close();
			//if(Conexion_Profit_1.State == ConnectionState.Open) Conexion_Profit_1.Close();
		}

		private Factura ResumenFactura(string idFactura,string idCliente,DateTime FechaE)
		{
			DataTable dtCxC = new DataTable();
			DataTable dtGiro = new DataTable();
            DataTable dtIncob = new DataTable();
			string str="";
			Factura Resumen;

			SQL=	"SELECT " +
						" Docum_cc.co_cli     AS CodClie, " +
						" Clientes.cli_des    AS Descrip, " +
						" Docum_cc.nro_doc    AS NroUnico, " +
						" Docum_cc.nro_orig   AS NroRegi, " +
						" Docum_cc.fec_emis   AS FechaE, " +
						" Docum_cc.fec_venc   AS FechaV, " +
						" Docum_cc.doc_orig   AS TipoCxc, " +
						" Docum_cc.monto_net  AS Monto, " +
						" Docum_cc.origen_d   AS NumeroD, " +
						" Docum_cc.saldo      AS Saldo, " +
						" Docum_cc.saldo      AS SaldoAct, " +
						" MIN(Docum_cc.monto_net) AS Giro, " + 
						" SUM(Docum_cc.monto_net) as Total, " +
						" SUM(Docum_cc.saldo) as SaldoT " +
					" FROM " +
						" Docum_cc INNER JOIN Clientes ON Docum_cc.co_cli = Clientes.co_cli " +
					" WHERE " +
                        " Docum_cc.co_cli = '" + idCliente + "' AND Docum_cc.fec_emis = '" + FechaE.ToString("yyyy-MM-dd")  + "' AND " +
                        " docum_cc.tipo_doc = 'GIRO' AND Docum_cc.fec_emis!=Docum_cc.fec_venc " +
					" GROUP BY " +
						" Docum_cc.co_cli, " +
						" Clientes.cli_des, " +
						" Docum_cc.nro_doc, " +
						" Docum_cc.nro_orig, " +
						" Docum_cc.fec_emis, " +
						" Docum_cc.fec_venc, " +
						" Docum_cc.doc_orig, " +
						" Docum_cc.monto_net, " +
						" Docum_cc.nro_doc, " +
						" Docum_cc.saldo, " +
						" Docum_cc.saldo, " +
                        " origen_d" +
					" ORDER BY " +
                        " Docum_cc.fec_venc ASC  ";

			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

			SQL=	"SELECT " +
				" min(Docum_cc.monto_net) AS Giro, SUM(Docum_cc.monto_net) as Total, SUM(Docum_cc.saldo) as Saldo " +
				" FROM " +
				" Docum_cc INNER JOIN Clientes ON Docum_cc.co_cli = Clientes.co_cli " +
				" WHERE " +
                " Docum_cc.co_cli = '" + idCliente + "' AND Docum_cc.fec_emis = '" + FechaE.ToString("yyyy-MM-dd") + "' AND docum_cc.tipo_doc = 'GIRO' ";

			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtGiro);

			Resumen = new Factura("");
			Resumen.Experiencia = 254;

			if(dtCxC.Rows.Count>0)
			{
				int Cuota=0;
				int Total=0;
				int Dias =0;
				double Monto=0.0;
				string strMonto;
				string cedula;
				bool experiencia = false;
                bool incobrable = false;
				DateTime fechaVenc;
				str = dtCxC.Rows[0]["NumeroD"].ToString().Trim();

				cedula = dtCxC.Rows[0]["CodClie"].ToString();

                if (idFactura == "") idFactura = str;

                if (idFactura.Length < 8)
                {
                  idFactura = idFactura.PadLeft(7, '0');
                  idFactura = "A" + idFactura;
                }

				Resumen = new Factura(idFactura);
				Resumen.Cliente = dtCxC.Rows[0]["Descrip"].ToString();
				Resumen.MontoTotal = 0.0;

				Resumen.PagoMensual = Convert.ToDouble(dtGiro.Rows[0]["Giro"].ToString());
				Resumen.MontoTotal = Convert.ToDouble(dtGiro.Rows[0]["Total"].ToString());
				Resumen.Saldo = Convert.ToDouble(dtGiro.Rows[0]["Saldo"].ToString());
				if(Convert.ToDouble(Resumen.Saldo)==0.0)
					Resumen.Cancelada = true;
				else Resumen.Cancelada = false;

				Total = dtCxC.Rows.Count;
                Resumen.Giros = dtCxC.Rows.Count;
				for(int i=0;i<dtCxC.Rows.Count;i++)
				{
					if(Convert.ToDouble(dtCxC.Rows[i]["Saldo"].ToString())>0.0)
					{
						Resumen.GirosSinCancelar++;
						Resumen.SaldoRestante = Convert.ToDouble(dtCxC.Rows[i]["Saldo"].ToString());

						fechaVenc = Convert.ToDateTime(dtCxC.Rows[i]["FechaV"].ToString());
						if(fechaVenc<DateTime.Now){
							Resumen.GirosVencidosSinCancelar++;
							Resumen.SaldoVencidoSinCancelar += Convert.ToDouble(dtCxC.Rows[i]["Saldo"].ToString());
						}

					}

					Cuota = i+1;

					DateTime FE = Convert.ToDateTime(dtCxC.Rows[i]["FechaE"].ToString());

                    Dias = Analizar_FechaV(dtCxC.Rows[i]["FechaV"].ToString(),_LaFecha);
					strMonto = dtCxC.Rows[i]["Saldo"].ToString();
					Monto = Convert.ToDouble(strMonto);

                    SQL = "select * from docum_cc where tipo_doc='AJNM' and nro_orig='' and observa like '%incobrable%' and co_cli='" + cedula + "'";
                    clsBD.EjecutarQuery(strConexion, Conexion, SQL, out dtIncob);

                    if (dtIncob.Rows.Count > 0) {
                        Dias = 121;
                        incobrable = true;
                    } 

					Resumen.Dias = Dias;
                    if (incobrable) {
                        Resumen.Experiencia = 21;
                    }
					else if((Monto!=0.0 || Dias == -1) && !experiencia)
					{
						if (Dias == -1)
							Resumen.Experiencia = 0;
						else if (Dias>=0 && Dias<=30)
							Resumen.Experiencia = 1;
						else if (Dias>30 && Dias<=60)
							Resumen.Experiencia = 2;
						else if (Dias>60 && Dias<=90)
							Resumen.Experiencia = 3;
						else if (Dias>90 && Dias<=120)
							Resumen.Experiencia = 4;
						else if (Dias>120)
							Resumen.Experiencia = 20;
						else if (Dias < 0 && FE.Month != _LaFecha.Month && FE.Year != _LaFecha.Year)
							Resumen.Experiencia = 1;
						experiencia=true;
					}
					else 
					{
						if(!experiencia)
						{
							Resumen.Experiencia = 1;

							if(Cuota==Total && Monto == 0.0 )
							{
								DataTable dtFechaC = new DataTable();
								string F = "";

								DateTime FV = Convert.ToDateTime(dtCxC.Rows[i]["FechaV"].ToString());

								SQL= " SELECT " +
									"cobros.fec_cob  " +
									" FROM " +
									"reng_cob INNER JOIN cobros ON reng_cob.cob_num = cobros.cob_num " +
									" WHERE " +
									"reng_cob.doc_num = '" + dtCxC.Rows[i]["NroUnico"].ToString() + "' " +
									"ORDER BY " +
									"cobros.fec_cob DESC ";

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
						}
					}
					
				}
				return Resumen;
			}
			else return Resumen;
		}


		public Documentos[] DocumentosSaint(System.Windows.Forms.ListBox list)
		{

			DataTable dtFact = new DataTable();
			DataTable dtCxC = new DataTable();
			DataTable dtDetFact = new DataTable();

			DateTime[] fechascompra = new DateTime[20];
			int indexfc = 0;
			string ultimacedula = "";

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
				//" FacturasViejas.MontoV>0  AND (FacturasViejas.Cedula = '10128631') " +
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
				//" SAFACT.MtoFinanc >0 AND (SAFACT.CodClie = '10128631') " +
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
				// Nuevo: Para buscar los registros perdidos
				/*
				if(ultimacedula=="") ultimacedula = dtFact.Rows[i]["CodClie"].ToString();
				if(ultimacedula!=dtFact.Rows[i]["CodClie"].ToString())
				{
					FECHA = " AND DATEDIFF(dd,'981104',SAACXC.FechaE)<>0 " +

					fechascompra = new DateTime[20];
					indexfc=0;

					#region
					SQL = "SELECT " +
							" SAACXC.CodClie, " +
							" SACLIE.Descrip, " +
							" SAACXC.NroUnico, " +
							" SAACXC.NroRegi, " +
							" SAACXC.FechaE, " +
							" SAACXC.FechaV, " +
							" SAACXC.TipoCxc, " +
							" SAACXC.Monto, " +
							" SAACXC.NumeroD, " +
							" SAACXC.Saldo, " +
							" SAACXC.SaldoAct, " +
							" SAFACT.NumeroD, " +
							" SAFACT.NroCtrol, " +
							" SAFACT.FechaE " +
							" FROM " +
							" SAFACT RIGHT JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie) " +
							" ON ((SAFACT.CodClie=SACLIE.CodClie) AND DATEDIFF(dd,SAFACT.FechaE,SAACXC.FechaE)=0) " +
							" WHERE " +
							" SAACXC.TipoCxc='60' " +
							" AND " +
							" SAACXC.CodClie = '7302399' " + FECHAS +
							" ORDER BY " +
							" SAACXC.CodClie ASC ";
					#endregion
				
				}
				else{
					fechascompra[indexfc] = dtFact.Rows[i]["FechaE"];
					indexfc++;
				}
				*/

				#region SQL CxC
				SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
					" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
					" WHERE SAACXC.TipoCxc='60' AND DATEDIFF(dd,'" + Convert.ToDateTime(dtFact.Rows[i]["FechaE"].ToString()).ToString("yyyy-MM-dd") + "',SAACXC.FechaE)=0 AND SAACXC.CodClie = '" + dtFact.Rows[i]["CodClie"].ToString() + "' " +
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


		public void MigrarProfit(Documentos[] doc,System.Windows.Forms.ListBox list)
		{
			string co_ve = "01";
			string co_us = "002";
			string co_su = "01";
			DateTime fechaE = new DateTime();
			DateTime fechaV = new DateTime();
			DataTable dtArt = new DataTable();

			int[] nro_doc_GIRO;

			DataTable dt = new DataTable();
			DataTable dtCli = new DataTable();

			int max_nro_doc_CFXG=0;
			int max_nro_doc_GIRO=0;
			int max_fact_num=0;
			int max_cobro_num=0;
			int max_num_mov_caj=1;

			double Monto_Cancelados = 0.0;
			int indice=0;

			Conexion_Profit.Open();

			SQL = "SELECT MAX(nro_doc) FROM docum_cc WHERE tipo_doc='CFXG';";
			clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dt);
			if(dt.Rows[0][0]!=DBNull.Value) max_nro_doc_CFXG = Convert.ToInt32(dt.Rows[0][0]);

			SQL = "SELECT MAX(nro_doc) FROM docum_cc WHERE tipo_doc='GIRO';";
			clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dt);
			if(dt.Rows[0][0]!=DBNull.Value) max_nro_doc_GIRO = Convert.ToInt32(dt.Rows[0][0]);

			SQL = "SELECT MAX(fact_num) FROM factura;";
			clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dt);
			if(dt.Rows[0][0]!=DBNull.Value) max_fact_num = Convert.ToInt32(dt.Rows[0][0]);

			SQL = "SELECT MAX(cob_num) FROM cobros;";
			clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dt);
			if(dt.Rows[0][0]!=DBNull.Value) max_cobro_num = Convert.ToInt32(dt.Rows[0][0]);

			max_fact_num += 1600;

			for(int i=0;i<doc.Length;i++)
			{
				if(!doc[i].fact.Cancelada)
				{
					try
					{
				
						#region Insertar Cliente

						// TODO: Verificar que el cliente existe

						SQL = "SELECT clientes.co_cli FROM clientes WHERE co_cli='" + doc[i].fact.IDCliente.PadRight(10) + "'";

						clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dtCli);

						if(dtCli.Rows.Count == 0)
						{

							SQL = "INSERT INTO " +
								"CLIENTES( " +
								"co_cli, " +
								"tipo, " +
								"cli_des, " +
								"direc1, " +
								"telefonos, " +
								"fecha_reg, " +
								"fec_ult_ve, " +
								"co_zon, " +
								"co_seg, " +
								"co_ven, " +
								"rif, " +
								"co_ingr, " +
								"co_us_in, " +
								"fe_us_in, " +
								"fe_us_mo, " +
								"fe_us_el, " +
								"co_sucu, " +
								"tipo_adi, " +
								"co_tab, " +
								"tipo_per, " +
								"estado, " +
								"co_pais, row_id " +
								") " +
								"VALUES( " +
								"'" + doc[i].fact.IDCliente.PadRight(10) + "', " +
								"'NAT', " +
								"'" + doc[i].fact.Cliente + "', " +
								"'" + doc[i].fact.Direccion + "', " +
								"'" + doc[i].fact.Telefono + "', " +
								"CAST('" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' AS DATETIME), " +
								"CAST('" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' AS DATETIME), " +
								"'01', " +
								"'01', " +
								"'" + co_ve + "', " +
								"'V-" + doc[i].fact.IDCliente.PadRight(10) + "', " +
								"'01', " +
								"'002', " +
								"CAST('" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' AS DATETIME), " +
								"CAST('" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' AS DATETIME), " +
								"CAST('" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' AS DATETIME), " +
								"'" + co_su + "', " +
								"1, " +
								"0, " +
								"1, " +
								"'A', " +
								"'VE    ', DEFAULT " +
								");";

							if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
							{
								list.Items.Add("Error => INSERT Cliente = " + doc[i].fact.IDCliente + " -> " + clsBD.MensajeError);
								list.Refresh();
								list.SendToBack();					
							}
						}

						#endregion

						#region Insertar docum_cc -> FACT
					
						max_fact_num++;

						SQL = "INSERT INTO [docum_cc] ([tipo_doc], [nro_doc], [anulado], [movi], [aut], [num_control], [co_cli], [contrib], [fec_emis], [fec_venc], [observa], [doc_orig], [nro_orig], [co_ban], [nro_che], [co_ven], [tipo], [tasa], [moneda], [monto_imp], [monto_gen], [monto_a1], [monto_a2], [monto_bru], [descuentos], [monto_des], [recargo], [monto_rec], [monto_otr], [monto_net], [saldo], [feccom], [numcom], [dis_cen], [comis1], [comis2], [comis3], [comis4], [adicional], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [revisado], [trasnfe], [numcon], [co_sucu], [mon_ilc], [otros1], [otros2], [otros3], [reng_si], [comis5], [comis6], [row_id], [aux01], [aux02], [salestax], [origen], [origen_d]) " +
							"VALUES " +
							" ('FACT', " + max_fact_num + ", 0, 0, 1, 0, '" + doc[i].fact.IDCliente.PadRight(10) + "', 1, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '" + doc[i].fact.FechaCancelacion.ToString("yyyy-MM-dd") + "', '', '', 0, '0', '', '" + co_ve + "', '1', 1, 'BS', " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.Impuesto)) + ", 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto)) + ", '', 0, '', 0, 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto + doc[i].fact.Impuesto)) + ", 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, ' " +
							" <IVA> " +
							" <1>14.000/" + Clase_ValidarDBL.ComaXPunto(doc[i].fact.MontoNeto.ToString("#########.00")) + "/" + Clase_ValidarDBL.ComaXPunto(doc[i].fact.Impuesto.ToString("#########.00")) + "</1> " +
							" </IVA>' " + 
							", 0, 0, 0, 0, 0, '', '', '', '', '', '', '', '', '" + co_us + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + co_us + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '', '', '" + co_su + "', 0, 0, 0, 0, 0, 0, 0, DEFAULT, 0, 'Luelher', '', '', '') ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT docum_cc = FACT = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						#endregion

						#region Insertar Cobro
						max_cobro_num++;

						SQL = "INSERT INTO [cobros] ([cob_num], [recibo], [co_cli], [co_ven], [fec_cob], [anulado], [monto], [dppago], [mont_ncr], [ncr], [tcomi_porc], [tcomi_line], [tcomi_art], [tcomi_conc], [feccom], [tasa], [moneda], [numcom], [dis_cen], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [recargo], [adel_num], [revisado], [trasnfe], [co_sucu], [descrip], [num_dev], [devdinero], [num_turno], [aux01], [aux02], [origen], [origen_d]) " +
							"VALUES " +
							"(" + max_cobro_num.ToString() + ", '', '" + doc[i].fact.IDCliente.PadRight(10) + "', '" + co_ve + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, 0, 0, 0, 0, 0, 0, 0, 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 1, 'BS', 0, '', '', '', '', '', '', '', '', '', '005', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', 0, 0, '', '', '" + co_su + "', '', 0, 0, 0, 0, 'Luelher', '', '') ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT Cobros = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();
						}

						#endregion

						#region Insertar docum_cc -> CFXG
						max_nro_doc_CFXG++;
						max_nro_doc_GIRO++;
						SQL = "INSERT INTO " +
							"Docum_cc( " +
							"co_cli, " +
							"tipo_doc, " +
							"nro_doc, " +
							"aut, " +
							"contrib, " +
							"fec_emis, " +
							"fec_venc, " +
							"observa, " +
							"doc_orig, " +
							"nro_orig, " + 
							"co_ven, " +
							"tipo, " +
							"tasa, " +
							"moneda, " +
							"monto_imp, " +
							"monto_bru, " +
							"recargo, " +
							"monto_net, " +
							"saldo, " +
							"feccom, " +
							"co_us_in, " +
							"fe_us_mo, " +
							"fe_us_el, " +
							"co_sucu, " +
							"row_id, " +
							"aux02" +
							") " +
							"VALUES(" +
							"'" + doc[i].fact.IDCliente.PadRight(10) + "', " +
							"'CFXG', " +
							"" + max_nro_doc_CFXG + ", " +
							"1, " +
							"1, " +
							"'" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', " +
							"'" + doc[i].fact.FechaCancelacion.ToString("yyyy-MM-dd") + "', " +
							"'GIROS " + max_nro_doc_GIRO.ToString() + "-" + Clase_ValidarDBL.ComaXPunto(Convert.ToString(max_nro_doc_GIRO+doc[i].fact.Giros)) + " FACT " + max_fact_num.ToString() + "', " + 
							"'COBR', " +
							"" + max_cobro_num.ToString() + ", " +
							"'" + co_ve + "', " +
							"6, " +
							"1, " +
							"'BS', " +
							"0, " +
							"" + Clase_ValidarDBL.ComaXPunto(Convert.ToDouble(doc[i].fact.MontoNeto+doc[i].fact.Impuesto).ToString("#########.00")) + ", " +
							"0, " +
							"" + Clase_ValidarDBL.ComaXPunto(Convert.ToDouble(doc[i].fact.MontoNeto+doc[i].fact.Impuesto).ToString("#########.00")) + ", " +
							"0, " +
							"'" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', " +
							"'" + co_us + "', " +
							"'" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', " +
							"'" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', " +
							"'" + co_su + "', " +
							"DEFAULT, " +
							"'Luelher' " +
							")";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT docum_cc = FACT = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						#endregion

						#region Factura
					
						string comentario="";
						for(int j=1;j<=doc[i].fact.Giros;j++)
						{
							comentario += "  #" + j + " Bs." + Clase_ValidarDBL.ComaXPunto(doc[i].fact.PagoMensual.ToString("#########.00"));
							if(j<doc[i].fact.Giros) comentario += ",";
						}

						string condicio = "";
						if((doc[i].fact.Giros+1)<10) condicio = "0"+Convert.ToString((doc[i].fact.Giros+1));
							else condicio = Convert.ToString((doc[i].fact.Giros+1));
						SQL = "INSERT INTO [factura] ([fact_num], [contrib], [nombre], [rif], [nit], [num_control], [status], [comentario], [descrip], [saldo], [fec_emis], [fec_venc], [co_cli], [co_ven], [co_tran], [dir_ent], [forma_pag], [tot_bruto], [tot_neto], [glob_desc], [tot_reca], [porc_gdesc], [porc_reca], [total_uc], [total_cp], [tot_flete], [monto_dev], [totklu], [anulada], [impresa], [iva], [iva_dev], [feccom], [numcom], [tasa], [moneda], [dis_cen], [vuelto], [seriales], [tasag], [tasag10], [tasag20], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [revisado], [trasnfe], [numcon], [co_sucu], [mon_ilc], [otros1], [otros2], [otros3], [num_turno], [aux01], [aux02], [ID], [salestax], [origen], [origen_d]) " +
							"VALUES " +
							"(" + max_fact_num.ToString() + ", 1, '', '', '', 0, '0', '" + comentario + "', '', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '" + doc[i].fact.FechaCancelacion.ToString("yyyy-MM-dd") + "', '" + doc[i].fact.IDCliente.PadRight(10) + "', '" + co_ve + "', '001', '', '" + condicio + "', " + Clase_ValidarDBL.ComaXPunto(doc[i].fact.MontoNeto.ToString()) + ", " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto+doc[i].fact.Impuesto)) + ", 0, 0, '', '', 0, 0, 0, 0, 0, 0, 1, " + Clase_ValidarDBL.ComaXPunto(doc[i].fact.Impuesto.ToString()) + ", 0, '19000101', 0, 1, 'BS', ' " + 
							"<IVA> " + 
							"<1>14.000/" + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto.ToString("#########.00"))) + "/" + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.Intereses.ToString("#########.00"))) + "</1> " +
							"</IVA>' " +
							", 0, 0, 14, 0, 0, '', '', '', '', '', '', '', '', '" + co_us + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '', '', '" + co_su + "', 0, 0, 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(doc[i].fact.MontoNeto.ToString("#########.00")) + ", '" + max_cobro_num.ToString() + "', -1, '', '', '') ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT factura = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}


						#endregion

						#region Articulos de la Factura

						for(int k=0;k<doc[i].detfact.Length;k++)
						{

							#region Insertar art

							SQL = "SELECT art.co_art FROM art WHERE co_art='"+doc[i].detfact[k].CodItem+"'";

							clsBD.EjecutarQuery(strConexion_Profit,Conexion_Profit,SQL,out dtArt);

							if(dtArt.Rows.Count == 0){
								SQL = "INSERT INTO [art] ([co_art], [art_des], [fecha_reg], [manj_ser], [co_lin], [co_cat], [co_subl], [co_color], [item], [ref], [modelo], [procedenci], [comentario], [co_prov], [ubicacion], [uni_venta], [uni_compra], [uni_relac], [relac_aut], [stock_act], [stock_com], [sstock_com], [stock_lle], [sstock_lle], [stock_des], [sstock_des], [suni_venta], [suni_compr], [suni_relac], [sstock_act], [relac_comp], [relac_vent], [pto_pedido], [stock_max], [stock_min], [prec_om], [prec_vta1], [fec_prec_v], [fec_prec_2], [prec_vta2], [fec_prec_3], [prec_vta3], [fec_prec_4], [prec_vta4], [fec_prec_5], [prec_vta5], [prec_agr1], [prec_agr2], [prec_agr3], [prec_agr4], [prec_agr5], [can_agr], [fec_des_p5], [fec_has_p5], [co_imp], [margen_max], [ult_cos_un], [fec_ult_co], [cos_pro_un], [fec_cos_pr], [cos_merc], [fec_cos_me], [cos_prov], [fec_cos_p2], [ult_cos_do], [fec_cos_do], [cos_un_an], [fec_cos_an], [ult_cos_om], [fec_ult_om], [cos_pro_om], [fec_pro_om], [tipo_cos], [mont_comi], [porc_cos], [mont_cos], [porc_gas], [mont_gas], [f_cost], [fisico], [punt_cli], [punt_pro], [dias_repos], [tipo], [alm_prin], [anulado], [tipo_imp], [dis_cen], [mon_ilc], [capacidad], [grado_al], [tipo_licor], [compuesto], [picture], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [revisado], [trasnfe], [co_sucu], [tuni_venta], [equi_uni1], [equi_uni2], [equi_uni3], [lote], [serialp], [valido], [atributo1], [vatributo1], [atributo2], [vatributo2], [atributo3], [vatributo3], [atributo4], [vatributo4], [atributo5], [vatributo5], [atributo6], [vatributo6], [garantia], [peso], [pie], [margen1], [margen2], [margen3], [margen4], [margen5], [imagen1], [imagen2], [i_art_des], [uni_emp], [rel_emp], [movil], [row_id]) " +
									" VALUES " +                                                                                                                                                                                                                                 //01 o UND
									" ('" + doc[i].detfact[k].CodItem + "', '" + doc[i].detfact[k].Descripcion + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '01', 'SAINT', '01', '01', '', '', 'MODSAINT', '01', '', '01', '', '01', '', 1, 1, -1, 0, 0, 0, 0, 0, 0, '01', '', 0, 0, 0, 0, 0, 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto/doc[i].detfact.Length)) + ", '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, 0, 0, 0, 0, 0, 1, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '', 0, 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto/doc[i].detfact.Length)) + " , '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 'ULCO', 0, 0, 0, 0, 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 0, 0, 0, 0, 'V', '', 0, '1', '', 0, 0, 0, '', 0, NULL, '', '', '', '', '', '', '', '', '" + co_us + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , '" + co_us + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', '', '', '" + co_su + "', 'UND', 1, 1, 1, 0, '', 0, 0, '', 0, '', 0, '', 0, '', 0, '', 0, '', '', 0, 0, 0, 0, 0, 0, 0, '', '', '', '', 1, 0, DEFAULT) ";

								if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
								{
									list.Items.Add("Error => INSERT art = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
									list.Refresh();
									list.SendToBack();					
								}							
							}


							#endregion

							#region reng_fac

							SQL = "INSERT INTO [reng_fac] ([fact_num], [reng_num], [dis_cen], [tipo_doc], [reng_doc], [num_doc], [co_art], [co_alma], [total_art], [stotal_art], [pendiente], [uni_venta], [prec_vta], [porc_desc], [tipo_imp], [isv], [reng_neto], [cos_pro_un], [ult_cos_un], [ult_cos_om], [cos_pro_om], [total_dev], [monto_dev], [prec_vta2], [anulado], [des_art], [seleccion], [cant_imp], [comentario], [total_uni], [mon_ilc], [otros], [nro_lote], [fec_lote], [pendiente2], [tipo_doc2], [reng_doc2], [num_doc2], [tipo_prec], [co_alma2], [aux01], [aux02]) " +
								" VALUES " +
								" (" + max_fact_num + ", " + (k+1) + ", '', '', 0, 0, '" + doc[i].detfact[k].CodItem + "', '01', 1, 0, 1, 'UND', " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto/doc[i].detfact.Length)) + ", '', '1', 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto/doc[i].detfact.Length)) + ", " + Clase_ValidarDBL.ComaXPunto(Convert.ToString((doc[i].fact.MontoNeto/doc[i].detfact.Length)/2)) + ", " + Clase_ValidarDBL.ComaXPunto(Convert.ToString((doc[i].fact.MontoNeto/doc[i].detfact.Length)/2)) + ", 0, 0, 0, 0, 0, 0, '', 0, 0, '', 1, 0, 0, '', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "' , 0, '', 0, 0, '1', '', 0, '') ";

							if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
							{
								list.Items.Add("Error => INSERT factura = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
								list.Refresh();
								list.SendToBack();					
							}

							#endregion

						}

						#endregion

						#region Giros (docum_cc)

						// TODO: se deben colocar todos los giros, se debe recordar que 
						// saint los borra los que han sido cancelados

						double Monto_Bru = 0.0; // Monto Cuota sin financiamiento
						double Monto_Otr = 0.0; // Monto a Recargar por financiamiento en los giros
						double Monto_Net = 0.0; // Monto total a pagar por giro Monto_Otr+Monto_Bru
						double Saldo = 0.0;		// Saldo de la cuota
						nro_doc_GIRO = new int[doc[i].fact.Giros+1];
						double descuento = 0.0;
						string porcentaje_descuento = "0";

						for(int j=1;j<=doc[i].fact.Giros;j++)
						{

							Monto_Bru = (doc[i].fact.MontoNeto + doc[i].fact.Impuesto)/doc[i].fact.Giros;
							Monto_Otr = doc[i].fact.Intereses/doc[i].fact.Giros;
							Monto_Net = doc[i].fact.PagoMensual;

							for(int t=0;t<doc[i].cxc.Length;t++)
							{
								if(doc[i].cxc[t].NroCuota == j)
								{
									Saldo = doc[i].cxc[t].Saldo;
									Monto_Net = doc[i].cxc[t].Monto;
									Monto_Otr = doc[i].fact.Intereses/doc[i].fact.Giros;
									descuento = 0.0;
									porcentaje_descuento = "0";
									fechaE = doc[i].cxc[t].FechaE;
									fechaV = doc[i].cxc[t].FechaV;
									break;
								}
								else 
								{
									Saldo=0.0;
									Monto_Net = doc[i].cxc[t].Monto;
									descuento = 0;
									porcentaje_descuento = "0";
									Monto_Otr = doc[i].fact.Intereses/doc[i].fact.Giros;;
									fechaE = DateTime.Now;
									fechaV = DateTime.Now;
								}
							}
						
							max_nro_doc_GIRO++;

							SQL = "INSERT INTO [docum_cc] ([tipo_doc], [nro_doc], [anulado], [movi], [aut], [num_control], [co_cli], [contrib], [fec_emis], [fec_venc], [observa], [doc_orig], [nro_orig], [co_ban], [nro_che], [co_ven], [tipo], [tasa], [moneda], [monto_imp], [monto_gen], [monto_a1], [monto_a2], [monto_bru], [descuentos], [monto_des], [recargo], [monto_rec], [monto_otr], [monto_net], [saldo], [feccom], [numcom], [dis_cen], [comis1], [comis2], [comis3], [comis4], [adicional], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [revisado], [trasnfe], [numcon], [co_sucu], [mon_ilc], [otros1], [otros2], [otros3], [reng_si], [comis5], [comis6], [row_id], [aux01], [aux02], [salestax], [origen], [origen_d]) " +
								"VALUES " +
								" ('GIRO', " + max_nro_doc_GIRO + ", 0, 0, 1, 0, '" + doc[i].fact.IDCliente.PadRight(10) + "', 1, '" + fechaE.ToString("yyyy-MM-dd") + "', '" + fechaV.ToString("yyyy-MM-dd") + "', 'GIRO " + j + "/" + doc[i].fact.Giros +" ; FACT " + max_fact_num + "' , 'CFXG', " + max_nro_doc_CFXG + ", '0', '', '" + co_ve + "', '6', 1, 'BS', 0, 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(Monto_Bru.ToString()) + ", '" + porcentaje_descuento + "', " + Clase_ValidarDBL.ComaXPunto(descuento.ToString()) + ", '0', 0, " + Clase_ValidarDBL.ComaXPunto(Monto_Otr.ToString()) + ", " + Clase_ValidarDBL.ComaXPunto(Monto_Net.ToString()) + ", " + Clase_ValidarDBL.ComaXPunto(Saldo.ToString()) + ", '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, ' " +
								" <IVA> " +
								" <E>" + Clase_ValidarDBL.ComaXPunto(Monto_Net.ToString("#########.00")) + "</E> " +
								" </IVA>' " + 
								", 0, 0, 0, 0, 0, '', '', '', '', '', '', '', '', '" + co_us + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '" + co_us + "', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '', '', '" + co_su + "', 0, 0, 0, 0, 0, 0, 0, DEFAULT, 0, 'Luelher', '', '', '') ";

							nro_doc_GIRO[j] = max_nro_doc_GIRO;

							if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
							{
								list.Items.Add("Error => INSERT docum_cc = GIRO = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
								list.Refresh();
								list.SendToBack();					
							}
						}

						#endregion

						#region reng_tip

						//SQL = "INSERT INTO [dbo].[reng_tip] ([cob_num], [reng_num], [tip_cob], [movi], [num_doc], [mont_doc], [mont_tmp], [moneda], [banco], [cod_caja], [des_caja], [fec_cheq], [nombre_ban], [numero], [devuelto], [operador], [clave]) " +
						//	"VALUES  " +
						//	"(" + max_cobro_num.ToString() + ", " + max_cobro_num.ToString() + ", 'EFEC', 0, '', 0, 0, 'BS', '', '', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '', 0, '', '') ";

						SQL = "INSERT INTO reng_tip " +
							" (cob_num,reng_num,tip_cob,num_doc,mont_doc,mont_tmp,banco,cod_caja,des_caja,fec_cheq, moneda) VALUES " +
							" (" + max_cobro_num.ToString() + ", 1, 'EFEC', '                    ', 0, 0, '      ', '01    ', '                    ', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 'BS      ')";


						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT reng_tip = " + max_cobro_num.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						#endregion

						#region reng_cob

						SQL = "INSERT INTO [reng_cob] ([cob_num], [reng_num], [tp_doc_cob], [doc_num], [neto], [neto_tmp], [dppago], [dppago_tmp], [reng_ncr], [co_ven], [comis1], [comis2], [comis3], [comis4], [sign_aju_c], [porc_aju_c], [por_cob], [comi_cob], [mont_cob], [sino_pago], [sino_reten], [monto_dppago], [monto_reten], [imp_pago], [monto_obj], [isv], [nro_fact], [moneda], [tasa], [numcon], [sustraen], [co_islr], [fec_emis], [fec_venc], [comis5], [comis6], [fact_iva], [ret_iva], [porc_retn], [porc_desc], [aux01], [aux02]) " +
							"VALUES  " +
							"(" + max_cobro_num + ", 1, 'FACT', " + max_fact_num + ", " + Clase_ValidarDBL.ComaXPunto(Convert.ToString((doc[i].fact.MontoNeto + doc[i].fact.Impuesto))) + ", 0, 0, 0, 0, '', 0, 0, 0, 0, '', 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto)) + ", 0, 0, 0, 0, 0, 0, 0, '', 'BS', 1, '', 0, '', '', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', 0, 0, 0, 0, 0, 0, 0, '') ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT reng_cob = " + max_cobro_num.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						SQL = "INSERT INTO [reng_cob] ([cob_num], [reng_num], [tp_doc_cob], [doc_num], [neto], [neto_tmp], [dppago], [dppago_tmp], [reng_ncr], [co_ven], [comis1], [comis2], [comis3], [comis4], [sign_aju_c], [porc_aju_c], [por_cob], [comi_cob], [mont_cob], [sino_pago], [sino_reten], [monto_dppago], [monto_reten], [imp_pago], [monto_obj], [isv], [nro_fact], [moneda], [tasa], [numcon], [sustraen], [co_islr], [fec_emis], [fec_venc], [comis5], [comis6], [fact_iva], [ret_iva], [porc_retn], [porc_desc], [aux01], [aux02]) " +
							"VALUES  " +
							"(" + max_cobro_num + ", 2, 'CFXG', " + max_nro_doc_CFXG + ", " + Clase_ValidarDBL.ComaXPunto(Convert.ToString((doc[i].fact.MontoNeto + doc[i].fact.Impuesto))) + ", 0, 0, 0, 0, '', 0, 0, 0, 0, '', 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(Convert.ToString(doc[i].fact.MontoNeto)) + ", 0, 0, 0, 0, 0, 0, 0, '', 'BS', 1, '', 0, '', '', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', 0, 0, 0, 0, 0, 0, 0, '') ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT reng_cob = " + max_cobro_num.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

					
						#endregion

						list.TopIndex = list.Items.Count-1;
					
						#region Insertar Cobro de Giros Cancelados

						Monto_Cancelados = 0.0;

						for(int t=0;t<doc[i].cxc.Length;t++)
						{
							if(doc[i].cxc[t].Cancelada) Monto_Cancelados += doc[i].cxc[t].Monto; 
							else if(!doc[i].cxc[t].Cancelada && doc[i].cxc[t].Saldo > 0 && doc[i].cxc[t].Saldo < doc[i].cxc[t].Monto) Monto_Cancelados += (doc[i].cxc[t].Monto - doc[i].cxc[t].Saldo);
						}

						if(Monto_Cancelados>0.0)
						{
							max_cobro_num++;

							//SQL = "INSERT INTO [cobros] ([cob_num], [recibo], [co_cli], [co_ven], [fec_cob], [anulado], [monto], [dppago], [mont_ncr], [ncr], [tcomi_porc], [tcomi_line], [tcomi_art], [tcomi_conc], [feccom], [tasa], [moneda], [numcom], [dis_cen], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [co_us_el], [fe_us_el], [recargo], [adel_num], [revisado], [trasnfe], [co_sucu], [descrip], [num_dev], [devdinero], [num_turno], [aux01], [aux02], [origen], [origen_d]) " +
							//	"VALUES " +
							//	"(" + max_cobro_num.ToString() + ", '', '" + doc[i].fact.IDCliente.PadRight(10) + "', '" + co_ve + "', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 0, " + Clase_ValidarDBL.ComaXPunto(Monto_Cancelados.ToString()) + ", 0, 0, 0, 0, 0, 0, 0, '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 1, 'BS', 0, '', '', '', '', '', '', '', '', '', '005', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', 0, 0, '', '', '" + co_su + "', '', 0, 0, 0, 0, 'Luelher', '', '') ";
						
							SQL = "INSERT INTO cobros " +
								" (cob_num,co_cli,co_ven,fec_cob,monto,feccom,tasa,moneda,co_us_in,fe_us_in,co_us_mo,fe_us_mo,co_us_el,fe_us_el,revisado,trasnfe,co_sucu) VALUES " +
								" (" + max_cobro_num.ToString() + ", '" + doc[i].fact.IDCliente.PadRight(10) + "', '01    ', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', " + Clase_ValidarDBL.ComaXPunto(Monto_Cancelados.ToString()) + ", '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', 1.00000, 'BS    ', '002   ', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '      ', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', '      ', '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "', ' ', 'L', '01    ')";

							if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
							{
								list.Items.Add("Error => INSERT Cobros -> cob_num = " + max_cobro_num.ToString() + " = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
								list.Refresh();
								list.SendToBack();

							}
						
							SQL = "INSERT INTO reng_tip " +
								" (cob_num,reng_num,tip_cob,num_doc,mont_doc,mont_tmp,banco,cod_caja,des_caja,fec_cheq) VALUES " +
								" (" + max_cobro_num.ToString() + ", 1, 'EFEC', '                    ', " + Clase_ValidarDBL.ComaXPunto(Monto_Cancelados.ToString()) + ", " + Clase_ValidarDBL.ComaXPunto(Monto_Cancelados.ToString()) + ", '      ', '01    ', 'Caja Saint          ', '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "')";

							if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
							{
								list.Items.Add("Error => INSERT reng_tip = " + doc[i].fact.IDCliente.PadRight(10) + " -> " + clsBD.MensajeError);
								list.Refresh();
								list.SendToBack();

							}

						}

						#endregion

						#region reng_cob de Giros Cancelados

						if(Monto_Cancelados>0.0)
						{
							indice = 1;
							double elmonto = 0.0;
							for(int t=0;t<doc[i].cxc.Length;t++)
							{
								if(doc[i].cxc[t].Cancelada || (!doc[i].cxc[t].Cancelada && doc[i].cxc[t].Saldo < doc[i].cxc[t].Monto))
								{
									//SQL = "INSERT INTO [reng_cob] ([cob_num], [reng_num], [tp_doc_cob], [doc_num], [neto], [neto_tmp], [dppago], [dppago_tmp], [reng_ncr], [co_ven], [comis1], [comis2], [comis3], [comis4], [sign_aju_c], [porc_aju_c], [por_cob], [comi_cob], [mont_cob], [sino_pago], [sino_reten], [monto_dppago], [monto_reten], [imp_pago], [monto_obj], [isv], [nro_fact], [moneda], [tasa], [numcon], [sustraen], [co_islr], [fec_emis], [fec_venc], [comis5], [comis6], [fact_iva], [ret_iva], [porc_retn], [porc_desc], [aux01], [aux02]) " +
									//	"VALUES  " +
									//	"(" + max_cobro_num + ", " + indice +", 'GIRO', " + nro_doc_GIRO[t+1] + ", 0, 0, 0, 0, 0, '', 0, 0, 0, 0, '', 0, 0, 0, " + Clase_ValidarDBL.ComaXPunto(doc[i].cxc[t].Monto.ToString()) + ", 0, 0, 0, 0, 0, 0, 0, '', 'BS', 1, '', 0, '', '', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', 0, 0, 0, 0, 0, 0, 0, '') ";
									if((!doc[i].cxc[t].Cancelada && doc[i].cxc[t].Saldo < doc[i].cxc[t].Monto)){
										elmonto = doc[i].cxc[t].Monto - doc[i].cxc[t].Saldo;
									}else elmonto = doc[i].cxc[t].Monto;
								
									SQL = "INSERT INTO reng_cob " +
										" (cob_num,reng_num,tp_doc_cob,doc_num,neto,neto_tmp,dppago,mont_cob,imp_pago,monto_obj,isv,moneda,tasa,fec_emis,fec_venc) VALUES " +
										" (" + max_cobro_num + ", " + indice +", 'GIRO', " + nro_doc_GIRO[t+1] + ", " + Clase_ValidarDBL.ComaXPunto(elmonto.ToString()) + ", 0.00, 0.00, " + Clase_ValidarDBL.ComaXPunto(elmonto.ToString()) + ", 0.00, 0.00, 0.00, 'BS        ', 1.00000, '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "') ";
									if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
									{
										list.Items.Add("Error => INSERT reng_cob = " + max_cobro_num.ToString() + " -> " + clsBD.MensajeError);
										list.Refresh();
										list.SendToBack();					
									}
									indice++;
								}
							}
						}
					
						#endregion

						#region mov_caj
						max_num_mov_caj++;

						SQL = "INSERT INTO mov_caj " +
							" (mov_num,codigo,cta_egre, tasa, moneda, banc_tarj,fe_us_in,fe_us_mo,fe_us_el,fecha,fecha_che,feccom,co_sucu) values " +
							" (" + max_num_mov_caj + "      ,'01'  ,'01'    ,1.00000,'BS' ,''        ,'" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "','01') ";
					
						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => INSERT mov_caj = " + max_num_mov_caj.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						SQL = "UPDATE mov_caj SET  " +
							" dep_num = 0, " +
							" dep_con = 0,  " +
							" origen = 'COB',  " +
							" DESCRIP = 'Migración: " + doc[i].fact.Cliente + " ',  " +
							" forma_pag = 'EF',  " +
							" codigo = '01',  " +
							" anulado     = 0,  " +
							" monto_h = " + Clase_ValidarDBL.ComaXPunto(Monto_Cancelados.ToString()) + ",  " +
							" tipo_op = 'I',  " +
							" ori_dep = 1,  " +
							" fecha = '" + doc[i].fact.FechaE.ToString("yyyy-MM-dd") + "', " + 
							" fecha_che = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',  " +
							" cob_pag = " + max_cobro_num.ToString() + ",  " +
							" tasa = 1.00000,  " +
							" moneda = 'BS',  " +
							" CO_US_IN = '002',  " +
							" CO_US_MO = '',  " +
							" CO_US_EL = '',  " +
							" FE_US_IN = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',  " +
							" FE_US_MO = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',  " +
							" FE_US_EL = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' ,  " +
							" feccom = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' ,  " +
							" co_sucu = '01'  " +
							" WHERE  " +
							" mov_num = " + max_num_mov_caj + " ";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => UPDATE mov_caj = " + max_num_mov_caj.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						#endregion

						#region Otros Cambios

						//SQL = "UPDATE  cajas SET saldo_a = saldo_a + 240.90000, saldo_e = saldo_e + 240.90000 WHERE cod_caja = '01'";

						SQL = "UPDATE reng_tip SET movi=" + max_num_mov_caj + " WHERE cob_num=" + max_cobro_num + " AND reng_num=1 AND movi=0";

						if(clsBD.EjecutarNonQuery(strConexion_Profit,Conexion_Profit,SQL)==-1)
						{
							list.Items.Add("Error => UPDATE reng_tip  = " + max_num_mov_caj.ToString() + " -> " + clsBD.MensajeError);
							list.Refresh();
							list.SendToBack();					
						}

						#endregion
						
						

						list.Items.Add("Cliente > " + doc[i].fact.IDCliente + " > Factura > " + doc[i].fact.NroFactura + " => Procedada..");
						//list.Items.Add("Procesados => " + i + " de " + doc.Length + " ");
						list.Refresh();
						list.SendToBack();		
					
					}
					catch (Exception ex)
					{
						list.Items.Add("ERROR GRAVE => Cliente = " + doc[i].fact.IDCliente + " => NO Procedada.. coderror = " + ex.Message);
						list.Refresh();
						list.SendToBack();					

					}
				}
				else{
					list.Items.Add("FACTURA CANCELADA => Cliente = " + doc[i].fact.IDCliente + " Factura => " + doc[i].fact.NroFactura + " => NO Procedada.. ");
					list.Refresh();
					list.SendToBack();					
				}

			}
			CerrarConexiones();

			string[] lineas = new string[list.Items.Count];

			for(int i=0;i<list.Items.Count;i++)
			{
				lineas[i] = list.Items[i].ToString();
			}

			SaveFileDialog SFD = new SaveFileDialog();
			SFD.ShowDialog();
			if(SFD.FileNames[0].Trim()!="")
			{
				try
				{
					ExportarLogs(lineas,SFD.FileNames[0]);
					Mensaje.Informar("Reporte Generado","Migrando a Profit");

				}
				catch(Exception ex)
				{Mensaje.Error(ex.Message,"Migrando a Profit");}
			}			

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
			else if((Tam==6 || Tam==7 || Tam==8) && De)
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
