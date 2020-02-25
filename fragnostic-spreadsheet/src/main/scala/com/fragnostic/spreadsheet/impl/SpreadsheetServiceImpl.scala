package com.fragnostic.spreadsheet.impl

import java.io.OutputStream
import java.util.{ Locale, UUID }

import better.files.{ File => BFFile, _ }
import com.fragnostic.spreadsheet.api.SpreadsheetServiceApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.slf4j.{ Logger, LoggerFactory }

import scala.util.{ Failure, Success, Try }

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
    private final val extensionXls: String = "xls"

    private def addHeader(wb: Workbook, header: Array[String], row: Row): Unit =
      header.zipWithIndex.foreach {
        case (head, idx) =>
          val cell = row.createCell(idx)
          cell.setCellValue(head)

          val style: CellStyle = wb.createCellStyle()
          style.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex)
          style.setFillPattern(FillPatternType.NO_FILL)

          val font: Font = wb.createFont()
          font.setColor(IndexedColors.GREY_80_PERCENT.getIndex)
          font.setBold(true)
          style.setFont(font)

          cell.setCellStyle(style)
      }

    private def getTmpPath: String = s"$tmpDir/${UUID.randomUUID().toString}.$extensionXls"

    private def write(os: OutputStream, bytes: Array[Byte]): Unit = os.write(bytes)

    override def save[T, S](
      locale: Locale,
      list: List[T],
      path: String,
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, String] =
      getWorkbook(locale, list, sheetName, headers, newRow) fold (error => Left(error),
        wb => save(wb, path) fold (
          error => Left(error),
          success => Right("spreadsheet.service.save.success")))

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
        success => Right(BFFile(tmpPath).loadBytes))
    }

    override def getWorkbook[T, S](
      locale: Locale,
      list: List[T],
      sheetName: String,
      headers: Array[String],
      newRow: (Locale, T, Row) => Row): Either[String, Workbook] =
      Try({
        val wb: Workbook = new HSSFWorkbook()
        val sheet: Sheet = wb.createSheet(sheetName)

        addHeader(wb, headers, sheet.createRow(0))

        list.zipWithIndex.foreach {
          case (entity, idx) => newRow(locale, entity, sheet.createRow(idx + 1))
        }

        headers.zipWithIndex.foreach {
          case (head, idx) => sheet.autoSizeColumn(idx)
        }

        Right(wb)
      }) getOrElse (Left("spreadsheet.service.get.workbook.error"))

    override def getWorkbook(bytes: Array[Byte]): Either[String, Workbook] = {
      val tmpPath: String = getTmpPath
      BFFile(tmpPath).outputStream.foreach(write(_, bytes))
      getWorkbook(tmpPath) fold (
        error => Left(error),
        wb => Right(wb))
    }

    override def getWorkbook(path: String): Either[String, Workbook] =
      if (path.endsWith(".xls")) {
        Try(new HSSFWorkbook(BFFile(path).newFileInputStream)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("spreadsheet.service.get.workbook.error.1")
          },
          wb => Right(wb))
      } else if (path.endsWith(".xlsx")) {
        Try(new XSSFWorkbook(BFFile(path).newFileInputStream)) fold (
          error => {
            logger.error(s"getWorkbook() - $error\n\tpath:$path")
            Left("spreadsheet.service.get.workbook.error.2")
          },
          wb => Right(wb))
      } else {
        Left("spreadsheet.service.get.workbook.error.wrong.extension")
      }

    override def save(workbook: Workbook, path: String): Either[String, String] =
      Try(
        for {
          fileOut <- BFFile(path).newFileOutputStream().autoClosed
        } yield {
          workbook.write(fileOut)
        }) match {
          case Success(_) => Right("spreadsheet.service.save.workbook.success")
          case Failure(exception) =>
            logger.error(s"save() - $exception")
            Left("spreadsheet.service.save.workbook.fail")
        }

    override def save(bytes: Array[Byte], path: String): Either[String, String] =
      Try(
        for {
          fos <- BFFile(path).newFileOutputStream().autoClosed
        } write(fos, bytes)) match {
          case Success(_) => Right("spreadsheet.service.save.bytes.success")
          case Failure(exception) =>
            logger.error(s"save() - $exception")
            Left("spreadsheet.service.save.bytes.fail")
        }

  }

}
