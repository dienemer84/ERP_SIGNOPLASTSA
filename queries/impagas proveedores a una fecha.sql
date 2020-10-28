SELECT 
  IF (f.`tipo`=1,"NC",(IF(f.`tipo`=2,"ND","FC"))) AS tipo,
  f.`NroFactura`, 
c.`razon`,   
  f.`FechaEmision`,
  r.`id`,
 IF (tipo=1,( f.`total_estatico` * f.cambio_a_patron) *-1, ( f.`total_estatico` * f.cambio_a_patron)) AS Importe
FROM  `AdminFacturas` f 
LEFT JOIN AdminRecibosDetalleFacturas df ON df.`idFactura` = f.`id` 
LEFT JOIN AdminRecibos r ON df.`idRecibo` = r.id  AND r.`fecha` <="2020-03-31"
LEFT JOIN clientes c ON f.`idCliente` = c.id 
WHERE f.`estado` = 2 AND f.`FechaEmision` <="2020-03-31" ORDER BY f.`FechaEmision` DESC