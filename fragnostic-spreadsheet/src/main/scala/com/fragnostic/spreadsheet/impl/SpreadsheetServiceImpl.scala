package com.fragnostic.spreadsheet.impl

import java.io._
import java.util.{ Locale, UUID }

import com.fragnostic.spreadsheet.api.SpreadsheetServiceApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.slf4j.{ Logger, LoggerFactory }

import scala.util.Try

/**
 * Created by fernandobrule on 5/19/17.
 *
 * https://poi.apache.org/spreadsheet/quick-guide.html
 *
 */
trait SpreadsheetServiceImpl extends SpreadsheetServiceApi {

  private[this] val logger: Logger = LoggerFactory.getLogger(getClass.getName)

  def spreadsheetService = new DefaultSpreadsheetService()

  class DefaultSpreadsheetService extends SpreadsheetServiceApi {

    private final val tmpDir: String = System.getProperty("java.io.tmpdir")

    private def getTmpPath: String = {
      val name: String = UUID.randomUUID().toString
      val extension: String = "xls"
      s"$tmpDir/$name.$extension"
    }

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
      basePath: String,
      fileName: String,
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, String] =
      getWorkbook(locale, list, sheetName, headers, newRow) fold (error => Left(error),
        wb => {

          val uuid = UUID.randomUUID().toString
          val path = s"$basePath/$fileName-$uuid.xls"

          save(wb, path) fold (
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
        error => Left("spreadsheet.service.get.bytes.error"),
        wb => getBytes(wb) fold (
          error => Left("spreadsheet.service.get.bytes.error"),
          wb => Right(wb)))

    override def getBytes(workbook: Workbook): Either[String, Array[Byte]] = {
      /*
      this does not work
      {
        val baos: ByteArrayOutputStream = new ByteArrayOutputStream()
        workbook.write(baos)
        baos.close()
        Right(baos.toByteArray)
      }
      --------------------------------
      A workarround will be used
      --------------------------------
      */
      val tmpPath: String = getTmpPath
      save(workbook, tmpPath) fold (
        error => Left("spreadsheet.service.get.bytes.error"),
        success => {
          import better.files.{ File => BFFile }
          Right(BFFile(tmpPath).loadBytes)
        })
    }

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
        case e: Exception => Left("spreadsheet.service.get.workbook.error")
        case _: Throwable => Left("spreadsheet.service.get.workbook.error")
      }

    override def getWorkbook(bytes: Array[Byte]): Either[String, Workbook] = {
      val tmpPath: String = getTmpPath
      val file: File = new File(tmpPath)
      val os: OutputStream = new FileOutputStream(file)
      os.write(bytes)
      os.close()

      getWorkbook(tmpPath) fold (
        error => Left(error),
        wb => Right(wb))

    }

    override def getWorkbook(path: String): Either[String, Workbook] = {
      val fis: FileInputStream = new FileInputStream(new File(path))
      if (path.endsWith(".xls")) {
        Try(new HSSFWorkbook(fis)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("spreadsheet.service.get.workbook.error.1")
          },
          wb => Right(wb))
      } else if (path.endsWith(".xlsx")) {
        Try(new XSSFWorkbook(fis)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("spreadsheet.service.get.workbook.error.2")
          },
          wb => Right(wb))
      } else {
        Left("spreadsheet.service.get.workbook.error.wrong.extension")
      }
    }

    override def save(workbook: Workbook, path: String): Either[String, String] = {
      val fileOut: FileOutputStream = new FileOutputStream(path)
      workbook.write(fileOut)
      fileOut.close()
      Right("spreadsheet.service.save.workbook.success")
    }

    override def save(bytes: Array[Byte], path: String): Either[String, String] = {
      val fos: FileOutputStream = new FileOutputStream(path)
      fos.write(bytes)
      fos.close()
      Right("spreadsheet.service.save.bytes.success")
    }

  }

}
