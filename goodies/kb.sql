UPDATE AdminRecibos SET fecha=fechaCreacion WHERE fecha IS NULL 

DELETE  FROM AdminRecibos WHERE fecha< '2008-01-01' OR fechaCreacion < '2008-01-01'

SELECT * FROM AdminRecibosDetalleRetenciones rr LEFT JOIN AdminRecibos r ON r.id = rr.idRecibo WHERE r.id IS NULL

SELECT * FROM AdminRecibos WHERE id IN (
	SELECT rr.idRecibo FROM AdminRecibosDetalleRetenciones rr LEFT JOIN AdminRecibos r ON r.id = rr.idRecibo
	WHERE r.id IS NULL
)

DELETE rr.* FROM AdminRecibosDetalleRetenciones rr LEFT JOIN AdminRecibos r ON r.id = rr.idRecibo WHERE r.id IS NULL


DELETE df.* FROM AdminRecibosDetalleFacturas df LEFT JOIN AdminRecibos r ON r.id = df.idRecibo WHERE r.id IS NULL

DELETE rd.* FROM AdminRecibosDepositos rd LEFT JOIN AdminRecibos r ON r.id = rd.idRecibo WHERE r.id IS NULL
DELETE rc.* FROM AdminRecibosCheques rc LEFT JOIN AdminRecibos r ON r.id = rc.idRecibo WHERE r.id IS NULL



SELECT * FROM Cheques c LEFT JOIN AdminRecibosCheques rc ON rc.idCheque = c.id WHERE rc.id IS NULL AND c.propio = 0 AND c.en_cartera = 0








