SELECT
  COUNT(idTiemposProcesos),
  idTiemposProcesos
FROM PlaneamientoTiemposProcesosDetalle
GROUP BY idTiemposProcesos
HAVING COUNT(idTiemposProcesos) > 1
ORDER BY COUNT(idTiemposProcesos) DESC

SELECT idTiemposProcesos, inico, fin, ((TIME_TO_SEC(TIMEDIFF(fin,inico)) / 60) / 60) AS dur_horas FROM PlaneamientoTiemposProcesosDetalle WHERE idTiemposProcesos = 91614

SELECT * FROM PlaneamientoTiemposProcesos WHERE id = 91614


SELECT
  SUM((TIME_TO_SEC(TIMEDIFF(ptpd.fin,ptpd.inico)) / 60) / 60) AS sum_horas,
  t.id_sector,
  ptp.codigoTarea
FROM PlaneamientoTiemposProcesos ptp
  INNER JOIN PlaneamientoTiemposProcesosDetalle ptpd
    ON ptpd.idTiemposProcesos = ptp.id
  INNER JOIN tareas t
    ON t.id = ptp.codigoTarea
WHERE ptp.idPedido = 1195
GROUP BY t.id_sector, ptp.codigoTarea
  