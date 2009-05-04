using System;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Diagnostics;
using GrupoEmporium.Datos;

namespace GrupoEmporium.Saint.Resportes
{
	/// <summary>
	/// Clase para generar reportes de Saint Enterprise.
	/// </summary>
	public class Saint
	{
		#region Variables

		string SQL="";
		DateTime _LaFecha;

		clsBDConexion strConexion = new clsBDConexion("agenciaroyal","enterpriseadminbd");
		//clsBDConexion strConexion = new clsBDConexion("agenciaroyal","enterpriseadminbd");
		
		//Este objeto conexion debe ser reemplazado al cambiar a SQL Server
		OleDbConnection Conexion;
		//OdbcConnection Conexion;

		private struct ResumenCxC
		{
			public string IDCliente;
			public string Cliente;

			public double Vencido;
			public int CuotasVencidas; 
			public int Cuotas; 
			public int CuotasxVencer;

			public double xVencer0_30;
			public double xVencer30_60;
			public double xVencer60_90;
			public double xVencer90_120;

			public ResumenCxC(string strIDCliente)
			{
				IDCliente=strIDCliente;
				Cliente="";

				Vencido=0.0;
				CuotasVencidas=0; 
				Cuotas=0; 
				CuotasxVencer=0;

				xVencer0_30=0.0;
				xVencer30_60=0.0;
				xVencer60_90=0.0;
				xVencer90_120=0.0;
			}

		}


		private struct FacturaCliente
		{
			public string IDCliente;
			public string Cliente;

			public string Telefono;

			public string NroFactura;
			public DateTime FechaE;
			public double MontoTotal;
			public double PagoMensual;
			public int Giros;
			public DateTime FechaCancelacion;
			private byte _Experiencia;

			public FacturaCliente(string strNroFactura)
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
				_Experiencia=0;
			}

			public byte Experiencia
			{
				get {return _Experiencia;}
				set {if (value > _Experiencia) _Experiencia=value;}
			}

		}


		#endregion

		#region Contructor
		public Saint()
		{
			//Conexion = new OleDbConnection("Provider=SQLOLEDB;" + strConexion.StringConexion);
			Conexion = new OleDbConnection(@"Provider=sqloledb;Data Source=agenciaroyal;Initial Catalog=enterpriseadminbd;User Id=aroyal\geuser;Password=r41gemporium;");
			_LaFecha = DateTime.Now;
		}
		public Saint(DateTime Fecha)
		{

			//Conexion = new OleDbConnection("Provider=SQLOLEDB;" + strConexion.StringConexion);
			Conexion = new OleDbConnection(@"Provider=sqloledb;Data Source=agenciaroyal;Initial Catalog=enterpriseadminbd;User Id=aroyal\geuser;Password=r41gemporium;");

			_LaFecha = Fecha;
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
			DataTable dt = new DataTable();
			DataTable dtReporte = new DataTable();
			DataRow dr;
			
			#region SQL Union Facturas

			SQL =	" SELECT SAFACT.CodClie, SACLIE.Descrip, SACLIE.Direc1, SACLIE.Direc2, SACLIE.Telef, SAFACT.NumeroD, SAFACT.NGiros, SAFACT.FechaE, SAFACT.FechaV, SAFACT.MtoFinanc " +
					" FROM SAFACT INNER JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) ON (SAFACT.CodClie=SACLIE.CodClie) AND (datevalue(SAFACT.FechaE) = datevalue(SAACXC.FechaE)) " +
					" WHERE SAFACT.FechaE<DateValue('" + _LaFecha.ToString("dd/MM/yy") + "') " +
					" AND " +
					" SAFACT.NGiros>0 " +
					" AND " +
					" SAFACT.MtoFinanc>0 " +
					" GROUP BY SAFACT.CodClie, SAFACT.NumeroD, SAFACT.NGiros, SAFACT.FechaE, SAFACT.FechaV, SAFACT.MtoFinanc, SACLIE.Descrip, SACLIE.Direc1, SACLIE.Direc2, SACLIE.Telef " +
					" ORDER BY SAFACT.NumeroD " +
					" UNION " +
					" SELECT FacturasViejas.Cedula AS CodClie, SACLIE.Descrip, SACLIE.Direc1, SACLIE.Direc2, SACLIE.Telef, FacturasViejas.NFactura AS NumeroD, FacturasViejas.GNumero AS NGiros, FacturasViejas.FechaE, FacturasViejas.FechaV, FacturasViejas.MontoV AS MtoFinanc " +
					" FROM FacturasViejas INNER JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) ON (FacturasViejas.Cedula=SACLIE.CodClie) AND (datevalue(FacturasViejas.FechaE) = datevalue(SAACXC.FechaE)) " +
					" WHERE FacturasViejas.FechaE<DateValue('" + _LaFecha.ToString("dd/MM/yy") + "') " +
					" AND " +
					" FacturasViejas.GNumero<>'0' " +
					" AND " +
					" FacturasViejas.MontoV>0 " +
					" GROUP BY FacturasViejas.Cedula, FacturasViejas.NFactura, FacturasViejas.GNumero, FacturasViejas.FechaE, FacturasViejas.FechaV, FacturasViejas.MontoV, SACLIE.Descrip, SACLIE.Direc1, SACLIE.Direc2, SACLIE.Telef; ";

			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

			#endregion

			if(dt.Rows.Count>0)
			{
				FacturaCliente RFactura;

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
					
					RFactura = ResumenFactura(dt.Rows[i]["Factura"].ToString(),dt.Rows[i]["Cedula"].ToString(),(DateTime)dt.Rows[i]["FechaE"]);
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


		public DataTable Reporte_Resumen_CXC()
		{
			DataTable dt = new DataTable();
			DataTable dtReporte = new DataTable();
			DataRow dr;
			
			SQL= "Select SAACXC.CodClie FROM SAACXC GROUP BY SAACXC.CodClie";
			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dt);

			Mensajes.Mensaje.Informar(dt.Rows.Count.ToString(),"Saint Reportes");

			if(dt.Rows.Count>0)
			{
				ResumenCxC CxCCliente;

				dtReporte.Columns.Add("Cedula",System.Type.GetType("System.String"));
				dtReporte.Columns.Add("Nombre",System.Type.GetType("System.String"));
				dtReporte.Columns.Add("Cuotas",System.Type.GetType("System.Int32"));
				dtReporte.Columns.Add("Vencidas",System.Type.GetType("System.Int32"));
				dtReporte.Columns.Add("TotalVencido",System.Type.GetType("System.Double"));
				dtReporte.Columns.Add("PorVencer",System.Type.GetType("System.Int32"));
				dtReporte.Columns.Add("Total0a30",System.Type.GetType("System.Double"));
				dtReporte.Columns.Add("Total30a60",System.Type.GetType("System.Double"));
				dtReporte.Columns.Add("Total60a90",System.Type.GetType("System.Double"));
				dtReporte.Columns.Add("Total90a120",System.Type.GetType("System.Double"));
				
				for(int i=0;i<dt.Rows.Count;i++)
				{
					
					CxCCliente = ResumenCxCCliente(dt.Rows[i][0].ToString());
					if(CxCCliente.IDCliente  !="")
					{
						dr = dtReporte.NewRow();
						dr["Cedula"] = CxCCliente.IDCliente;
						dr["Nombre"] = CxCCliente.Cliente;
						dr["Cuotas"] = CxCCliente.Cuotas;
						dr["Vencidas"] = CxCCliente.CuotasVencidas;
						dr["TotalVencido"] = CxCCliente.Vencido;
						dr["PorVencer"] = CxCCliente.CuotasxVencer;
						dr["Total0a30"] = CxCCliente.xVencer0_30;
						dr["Total30a60"] = CxCCliente.xVencer30_60;
						dr["Total60a90"] = CxCCliente.xVencer60_90;
						dr["Total90a120"] = CxCCliente.xVencer90_120;
						dtReporte.Rows.Add(dr);
						//Debug.WriteLine("Registro " + i.ToString());
					}
				}
				dtReporte.AcceptChanges();
				return dtReporte;
			}return new DataTable();

		}

		#endregion

		#region Metodos Privados

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
							}
						}
						return Resumen;
					}else return new ResumenCxC("");
				}
			}else return new ResumenCxC("");
		}

		private FacturaCliente ResumenFactura(string idFactura,string idCliente,DateTime FechaE)
		{
			DataTable dtCxC = new DataTable();
			string str="";
			FacturaCliente Resumen;

			SQL= "SELECT SAACXC.CodClie, SACLIE.Descrip, SAACXC.NroUnico, SAACXC.NroRegi, SAACXC.FechaE, SAACXC.FechaV, SAACXC.TipoCxc, SAACXC.Monto, SAACXC.NumeroD, SAACXC.Saldo, SAACXC.SaldoAct " +
				" FROM SAACXC INNER JOIN SACLIE ON SAACXC.CodClie = SACLIE.CodClie" + 
				" WHERE SAACXC.CodClie='" + idCliente + "' AND SAACXC.TipoCxc='60' AND datevalue(SAACXC.FechaE) = datevalue('" + FechaE.ToString("dd/MM/yy") + "')" +
				" ORDER BY SAACXC.NroUnico ASC; ";
			clsBD.EjecutarQuery(strConexion,Conexion,SQL,out dtCxC);

			if(dtCxC.Rows.Count>0)
			{
				if(dtCxC.Rows.Count>20)
				{
					str = dtCxC.Rows[0][0].ToString();

					return new FacturaCliente("");
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
						Resumen = new FacturaCliente(str = idFactura);
						Resumen.Cliente = dtCxC.Rows[0]["Descrip"].ToString();

						for(int i=0;i<dtCxC.Rows.Count;i++)
						{
							Dias = Analizar_FechaV(dtCxC.Rows[i]["FechaV"].ToString(),_LaFecha);
							strMonto = dtCxC.Rows[i]["Saldo"].ToString();
							Monto = Convert.ToDouble(strMonto);

							if(Monto!=0.0)
							{
								if (Dias>=0 && Dias<30)
									Resumen.Experiencia = 2;
								else if (Dias>=30 && Dias<60)
									Resumen.Experiencia = 3;
								else if (Dias>=60 && Dias<90)
									Resumen.Experiencia = 4;
								else if (Dias>=90 && Dias<120)
									Resumen.Experiencia = 5;
								else if (Dias>=120)
									Resumen.Experiencia = 20;
							}
							else Resumen.Experiencia = 1;
						}
						return Resumen;
					}
					else return new FacturaCliente("");
				}
			}
			else return new FacturaCliente("");
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

			TimeSpan Tiempo = Fecha.Subtract(FechaActual);

			if((int)Tiempo.TotalDays < 0) return -1;
			else return (int)Tiempo.TotalDays;
		}

		#endregion

	}
}
