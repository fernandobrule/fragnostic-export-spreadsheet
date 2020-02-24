package com.fragnostic.export.spreadsheet.impl

import java.io._
import java.util.{ Locale, UUID }

import com.fragnostic.export.spreadsheet.api.ExportSpreadsheetServiceApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.slf4j.LoggerFactory

import scala.util.Try

/**
 * Created by fernandobrule on 5/19/17.
 *
 * https://poi.apache.org/spreadsheet/quick-guide.html
 *
 */
trait ExportSpreadsheetServiceImpl extends ExportSpreadsheetServiceApi {

  private def logger = LoggerFactory.getLogger(getClass)

  def exportSpreadsheetService = new DefaultExportSpreadsheetService()

  class DefaultExportSpreadsheetService extends ExportSpreadsheetServiceApi {

    private def addHeader(wb: Workbook, header: Array[String], row: Row): Unit =
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

    override def save[T, S](
      locale: Locale,
      list: List[T],
      basePathExport: String,
      fileName: String,
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, String] =
      getBytes(locale, list, sheetName, headers, newRow) fold (error => Left(error),
        bytes => {

          val uuid = UUID.randomUUID().toString
          val path = s"$basePathExport/$fileName-$uuid.xls"

          save(bytes, path) fold (
            error => Left(error),
            success => Right(uuid))

        })

    override def getBytes[T, S](
      locale: Locale,
      list: List[T],
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, Array[Byte]] =
      getWorkbook(locale, list, sheetName, headers, newRow) fold (
        error => Left("export.spreadsheet.service.get.bytes.error"),
        wb => getBytes(wb) fold (
          error => Left("export.spreadsheet.service.get.bytes.error"),
          wb => Right(wb)))

    override def getBytes(workbook: Workbook): Either[String, Array[Byte]] = ???
    /*
    this does not work
    {
      val baos: ByteArrayOutputStream = new ByteArrayOutputStream()
      workbook.write(baos)
      baos.close()
      Right(baos.toByteArray)
    }
    */

    override def getWorkbook[T, S](
      locale: Locale,
      list: List[T],
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, Workbook] =
      try {

        val wb: Workbook = new HSSFWorkbook()
        val sheet: Sheet = wb.createSheet(sheetName)

        //
        // AGREGA HEADER
        if (logger.isInfoEnabled) logger.info(s"workbook() - headers.length:${headers.length}")
        addHeader(wb, headers, sheet.createRow(0))

        //
        // AGREGA FILAS
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
        case e: Exception => Left("export.spreadsheet.service.get.workbook.error")
        case _: Throwable => Left("export.spreadsheet.service.get.workbook.error")
      }

    override def getWorkbook(bytes: Array[Byte]): Either[String, Workbook] = {

      val os: OutputStream = new ByteArrayOutputStream(bytes.length)
      os.write(bytes, 0, bytes.length)

      val wb: Workbook = new HSSFWorkbook()
      wb.write(os)

      Right(wb)
    }

    override def getWorkbook(path: String): Either[String, Workbook] = {
      val fis: FileInputStream = new FileInputStream(new File(path))
      if (path.endsWith(".xls")) {
        Try(new HSSFWorkbook(fis)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("export.spreadsheet.service.get.workbook.error.1")
          },
          wb => Right(wb))
      } else if (path.endsWith(".xlsx")) {
        Try(new XSSFWorkbook(fis)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("export.spreadsheet.service.get.workbook.error.2")
          },
          wb => Right(wb))
      } else {
        Left("export.spreadsheet.service.get.workbook.error.wrong.extension")
      }
    }

    override def save(workbook: Workbook, path: String): Either[String, String] = {
      val fileOut: FileOutputStream = new FileOutputStream(path)
      workbook.write(fileOut)
      fileOut.close()
      Right("export.spreadsheet.service.save.workbook.success")
    }

    override def save(bytes: Array[Byte], path: String): Either[String, String] = {
      val fos: FileOutputStream = new FileOutputStream(path)
      fos.write(bytes)
      fos.close()
      Right("export.spreadsheet.service.save.bytes.success")
    }

  }

}
