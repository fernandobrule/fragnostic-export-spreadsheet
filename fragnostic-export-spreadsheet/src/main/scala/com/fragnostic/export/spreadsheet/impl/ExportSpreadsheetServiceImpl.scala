package com.fragnostic.export.spreadsheet.impl

import java.io.{ ByteArrayOutputStream, FileOutputStream }
import java.util.{ Locale, UUID }

import com.fragnostic.export.spreadsheet.api.ExportSpreadsheetServiceApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel._
import org.slf4j.LoggerFactory

/**
 * Created by fernandobrule on 5/19/17.
 *
 * https://poi.apache.org/spreadsheet/quick-guide.html
 *
 */
trait ExportSpreadsheetServiceImpl extends ExportSpreadsheetServiceApi {

  private def logger = LoggerFactory.getLogger(getClass)

  def exportSpreadsheetService = new DefaultExportSpreadsheetService()

  def getBytes(wb: Workbook): Array[Byte] = {

    val bos: ByteArrayOutputStream = new ByteArrayOutputStream()
    try {
      wb.write(bos)
    } finally {
      bos.close()
    }

    bos.toByteArray

  }

  //
  // HEADER
  def addHeader(wb: Workbook, header: Array[String], row: Row): Unit =
    header.zipWithIndex.foreach {
      case (head, idx) =>
        if (logger.isInfoEnabled) logger.info(s"addHeader() - col:$idx - head:$head")

        val cell = row.createCell(idx)
        cell.setCellValue(head)

        // CELL STYLE
        val style: CellStyle = wb.createCellStyle()
        style.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex)
        style.setFillPattern(FillPatternType.NO_FILL)

        // CELL FONT
        val font: Font = wb.createFont()
        font.setColor(IndexedColors.GREY_80_PERCENT.getIndex)
        font.setBold(true)
        style.setFont(font)

        cell.setCellStyle(style)
    }

  class DefaultExportSpreadsheetService extends ExportSpreadsheetServiceApi {

    override def spreadsheet[T, S](
      list: List[T],
      basePathExport: String,
      fileName: String,
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, String] =
      bytes(list, sheetName, headers, newRow) fold (error => Left(error),
        bytes => {
          val uuid = UUID.randomUUID().toString
          val ruta = s"$basePathExport/$fileName-$uuid.xls"
          try {
            val fos: FileOutputStream = new FileOutputStream(ruta)
            fos.write(bytes)
            fos.close()

            if (logger.isInfoEnabled) logger.info(s"export | bytes saved on ruta:$ruta")

            Right(uuid)

          } catch {
            case e: Exception =>
              logger.error(s"export | $e")
              Left("export.service.error")
            case e: Throwable =>
              logger.error(s"export | $e")
              Left("export.service.error")
          }
        })

    override def bytes[T, S](
      list: List[T],
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, Array[Byte]] =
      workbook(list, sheetName, headers, newRow) fold (error => Left("export.service.error"),
        wb => Right(getBytes(wb)))

    override def workbook[T, S](
      list: List[T],
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, Workbook] =
      try {

        val wb: Workbook = new HSSFWorkbook()
        val createHelper: CreationHelper = wb.getCreationHelper
        val sheet: Sheet = wb.createSheet(sheetName)

        //
        // AGREGA HEADER
        if (logger.isInfoEnabled) logger.info(s"workbook() - headers.length:${headers.length}")
        addHeader(wb, headers, sheet.createRow(0))

        //
        // AGREGA FILAS
        val locale = new Locale.Builder().setLanguage("es").setRegion("CL").build()
        list.zipWithIndex.foreach {
          case (entity, idx) => newRow(locale, entity, sheet.createRow(idx + 1))
        }

        //
        // AUTO AJUSTA ANCHO DE COLUMNAS
        headers.zipWithIndex.foreach {
          case (head, idx) => sheet.autoSizeColumn(idx)
        }

        Right(wb)

      } catch {
        case e: Exception => Left("export.service.error")
        case _: Throwable => Left("export.service.error")
      }
  }

}
