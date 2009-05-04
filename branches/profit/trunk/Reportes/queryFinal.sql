SELECT 
	FacturasViejas.Cedula AS CodClie, 
	SACLIE.Descrip, 
	SACLIE.Direc1, 
	SACLIE.Direc2, 
	SACLIE.Telef, 
	cast(FacturasViejas.NFactura as varchar(10)) AS NumeroD, 
	FacturasViejas.GNumero AS NGiros, 
	FacturasViejas.FechaE, 
	FacturasViejas.FechaV, 
	CAST(FacturasViejas.MontoV AS decimal(13)) AS MtoFinanc
FROM 
	FacturasViejas INNER JOIN (SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) 
	ON (FacturasViejas.Cedula=SACLIE.CodClie) AND ((FacturasViejas.FechaE) = (SAACXC.FechaE))  
WHERE 
	DateDiff(dd,FacturasViejas.FechaE,CAST('01/10/05' AS DATETIME)) > 0  AND  
	FacturasViejas.GNumero<>'0'  AND  
	FacturasViejas.MontoV>0  
GROUP BY 
	FacturasViejas.Cedula, 
	FacturasViejas.NFactura, 
	FacturasViejas.GNumero, 
	FacturasViejas.FechaE, 
	FacturasViejas.FechaV, 
	FacturasViejas.MontoV, 
	SACLIE.Descrip, 
	SACLIE.Direc1, 
	SACLIE.Direc2, 
	SACLIE.Telef
UNION
SELECT 
	SAFACT.CodClie, 
	SACLIE.Descrip, 
	SACLIE.Direc1, 
	SACLIE.Direc2, 
	SACLIE.Telef, 
	SAFACT.NumeroD, 
	SAFACT.NGiros, 
	SAFACT.FechaE, 
	SAFACT.FechaV, 
	SAFACT.MtoFinanc  
FROM 
	SAFACT INNER JOIN 
	(SAACXC INNER JOIN SACLIE ON SAACXC.CodClie=SACLIE.CodClie) ON 
	((SAFACT.CodClie=SACLIE.CodClie) AND DATEDIFF(dd,SAFACT.FechaE,SAACXC.FechaE)=0)
WHERE 
	DATEDIFF(dd,SAFACT.FechaE,CAST('01/11/05' AS DATETIME)) < 0 AND  
	SAFACT.NGiros>0  AND
	SAFACT.MtoFinanc >0
GROUP BY 
	SAFACT.CodClie, 
	SAFACT.NumeroD, 
	SAFACT.NGiros, 
	SAFACT.FechaE, 
	SAFACT.FechaV, 
	SAFACT.MtoFinanc, 
	SACLIE.Descrip, 
	SACLIE.Direc1, 
	SACLIE.Direc2, 
	SACLIE.Telef  
ORDER BY 
	SAFACT.NumeroD;