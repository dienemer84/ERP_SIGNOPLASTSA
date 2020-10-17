SELECT * FROM (

SELECT DISTINCT 'contado' AS tipo, fp.numero_factura, p.id AS cod_proveedor, p.razon, fp.fecha, fp.fecha_carga, opf.id_orden_pago FROM AdminComprasFacturasProveedores fp INNER JOIN proveedores p ON fp.id_proveedor=p.id LEFT JOIN ordenes_pago_facturas opf ON opf.id_factura_proveedor=fp.id
WHERE tipo_doc_contable=1 AND fecha >='2019-04-01' AND fecha<='2020-03-31' AND p.estado=2 

UNION ALL

SELECT DISTINCT 'cta_cte' AS tipo, fp.numero_factura, p.id AS cod_proveedor, p.razon, fp.fecha, fp.fecha_carga, opf.id_orden_pago FROM AdminComprasFacturasProveedores fp INNER JOIN proveedores p ON fp.id_proveedor=p.id LEFT JOIN ordenes_pago_facturas opf ON opf.id_factura_proveedor=fp.id
WHERE tipo_doc_contable=1 AND fecha >='2019-04-01' AND fecha<='2020-03-31' AND p.estado=1
) X ORDER BY x.tipo, x.fecha