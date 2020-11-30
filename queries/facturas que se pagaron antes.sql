  SELECT op.id AS idop, fp.id, fp.numero_factura,p.razon,op.fecha AS fecha_op,fp.fecha AS fecha_factrura, fp.forma_de_pago_cta_cte AS es_cta_cte
	
	
	FROM 	ordenes_pago op 
	LEFT JOIN ordenes_pago_facturas opf ON opf.id_orden_pago = op.id
	LEFT JOIN AdminComprasFacturasProveedores fp ON opf.id_factura_proveedor=fp.id
	LEFT JOIN proveedores p ON fp.id_proveedor=p.id


WHERE op.fecha < fp.fecha AND op.fecha>='2018-04-01' AND op.fecha<='2020-03-31'
 #AND fp.forma_de_pago_cta_cte=0