

CREATE PROCEDURE `sp_permisos`.`spProcesarPadrones`()

    BEGIN
		    
		    TRUNCATE sp_permisos.Padron_Detalles_Ant;
		INSERT INTO sp_permisos.Padron_Detalles_Ant  (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron, FechaDesdePercepcion, FechaHastaPercepcion, FechaDesdeRetencion, FechaHastaRetencion)
		(SELECT Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron, FechaDesdePercepcion, FechaHastaPercepcion, FechaDesdeRetencion, FechaHastaRetencion FROM sp_permisos.Padron_Detalles_Ant LIMIT 1000);


		TRUNCATE sp_permisos.Padron_Detalles;
		INSERT INTO sp_permisos.Padron_Detalles (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron, FechaDesdePercepcion, FechaHastaPercepcion, FechaDesdeRetencion, FechaHastaRetencion)

		SELECT * FROM
		(

				(SELECT p.Cuit, p.Alicuota AS AlicuotaPercepcion, IFNULL(r.Alicuota,'0,00') AS AlicuotaRetencion, '1', p.FechaDesde AS pd, p.FechaHasta AS ph,r.FechaDesde,r.FechaHasta 
						
					FROM sp_permisos.IIBB2_Percepcion p LEFT JOIN sp_permisos.IIBB2_Retencion r ON r.Cuit=p.Cuit LIMIT 1000)

				 UNION ALL
					(SELECT r.Cuit, r.Alicuota AS AlicuotaRetencion, 0, '1','' AS c,'' AS d, r.FechaDesde AS a, r.FechaHasta AS b 
					FROM sp_permisos.IIBB2_Retencion r 
					WHERE NOT EXISTS (SELECT * FROM sp_permisos.IIBB2_Percepcion p WHERE p.Cuit=r.Cuit LIMIT 1000)
					LIMIT 1000)
				  UNION ALL
					(SELECT c.Cuit, c.AlicuotaPercepcion, c.AlicuotaRetencion, '2' ,c.FechaDesde, c.FechaHasta, c.FechaDesde, c.FechaHasta 
					FROM sp_permisos.IIBB2_Padron_CABA c	LIMIT 1000)
		) p
    
    

    END;
