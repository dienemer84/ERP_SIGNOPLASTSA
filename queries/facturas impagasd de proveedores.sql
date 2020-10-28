
SELECT 
fp.id,
  IF (
    fp.tipo_doc_contable = 0,
    "FC",
    IF (
      fp.tipo_doc_contable = 1,
      "NC",
      "ND"
    )
  ) AS tipo,
  fp.`numero_factura`,
  fp.`fecha`,
  p.`razon`,
 
 
  IF (
    fp.tipo_doc_contable = 1,
    - 1 ,1 ) AS mult,
    
    IF (fp.tipo_cambio IS NULL,1,fp.`tipo_cambio`) AS tipo_cambio,
        fp.impuesto_interno,
        fp.redondeo_iva,
        fp.monto_neto,
        IF (SUM(DISTINCT i.valor *(a.`alicuota`/100)) IS NULL,0,SUM(DISTINCT i.valor *(a.`alicuota`/100))) AS iva,
        IF (SUM(DISTINCT pe.valor) IS NULL,0,SUM(DISTINCT pe.valor)) AS percep
FROM
  AdminComprasFacturasProveedores fp 
  INNER JOIN proveedores p 
    ON fp.`id_proveedor` = p.id 
  JOIN `AdminComprasFacturasProveedoresIva` i 
    ON i.`id_factura_proveedor` = fp.`id` 
  JOIN `AdminConfigIvaAlicuotas` a 
    ON i.id_iva = a.id
   LEFT JOIN `AdminComprasFacturasProveedoresPercepciones` pe ON pe.`id_factura_proveedor`=fp.`id` 
WHERE fp.estado != 1 
  AND fp.id NOT IN 
  (SELECT DISTINCT 
    id_factura_proveedor 
  FROM
    `ordenes_pago_facturas` 
  WHERE id_orden_pago IN 
    (SELECT DISTINCT 
      id 
    FROM
      ordenes_pago op 
    WHERE op.fecha <= "2020-03-31" 
      AND op.`estado` = 1)) 
  AND fp.`fecha` <= "2020-03-31" 
  

GROUP BY fp.id
  ORDER BY p.razon