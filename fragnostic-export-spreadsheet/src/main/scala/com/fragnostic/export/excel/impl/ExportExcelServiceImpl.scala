package com.fragnostic.export.excel.impl

import java.io.{ ByteArrayOutputStream, FileOutputStream }
import java.util.{ Locale, UUID }

import com.fragnostic.export.excel.api.ExportExcelServiceApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel._
import org.slf4j.LoggerFactory

/**
 * Created by fernandobrule on 5/19/17.
 *
 * https://poi.apache.org/spreadsheet/quick-guide.html
 *
 */
trait ExportExcelServiceImpl extends ExportExcelServiceApi {

  private def logger = LoggerFactory.getLogger(getClass)

  private val debug = false

  def exportExcelService = new DefaultExportExcelService()

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
  def addHeader(wb: Workbook, header: Array[String], row: Row) =
    header.zipWithIndex.foreach {
      case (head, idx) =>

        if (debug) logger.info(s"col:$idx - head:$head")

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

  class DefaultExportExcelService extends ExportExcelServiceApi {

    override def export[T, S](list: List[T], fileName: String, sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, String] =
      exportBytes(list, sheetName, headers, enables, newRow) fold (
        error => Left(error),
        bytes => {

          val uuid = UUID.randomUUID().toString

          val javaIoTmpdir = System.getProperty("java.io.tmpdir")
          val basePath = javaIoTmpdir // "/Users/fernandobrule/Tmp"

          val ruta = s"$basePath/$fileName-$uuid.xls"
          try {
            val fos: FileOutputStream = new FileOutputStream(ruta)
            fos.write(bytes)
            fos.close()

            if (debug) logger.info(s"export | bytes saved on ruta:$ruta")

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

    override def exportBytes[T, S](list: List[T], sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, Array[Byte]] =
      exportWb(list, sheetName, headers, enables, newRow) fold (
        error => Left("export.service.error"),
        wb => Right(getBytes(wb)))

    override def exportWb[T, S](list: List[T], sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, Workbook] =
      try {

        val wb: Workbook = new HSSFWorkbook()
        val createHelper: CreationHelper = wb.getCreationHelper
        val sheet: Sheet = wb.createSheet(sheetName)

        //
        // AGREGA HEADER
        addHeader(wb, headers, sheet.createRow(0))

        //
        // AGREGA FILAS
        val locale = new Locale.Builder().setLanguage("es").setRegion("CL").build()
        list.zipWithIndex.foreach {
          case (entity, idx) => newRow(locale, entity, sheet.createRow(idx + 1), enables)
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
