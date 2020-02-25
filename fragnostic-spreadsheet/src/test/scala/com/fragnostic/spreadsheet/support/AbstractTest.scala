package com.fragnostic.spreadsheet.support

import java.io.File
import java.util.Locale

import org.apache.poi.ss.usermodel.{ Cell, CellStyle, Row }
import org.scalatest.{ FunSpec, Matchers }
import better.files.{ Resource => BFResource }
import scala.util.Random
/**
 * Created by fernandobrule on 5/20/17.
 */
abstract class AbstractTest extends FunSpec with Matchers {

  lazy val locale: Locale = new Locale.Builder().setRegion("CL").setLanguage("es").build

  lazy val fileNameXls: String = "the-spreadsheet.xls"
  lazy val fileNameXlsx: String = "the-spreadsheet.xlsx"

  lazy val pathXls: String = BFResource.getUrl(fileNameXls).getPath
  lazy val pathXlsx: String = BFResource.getUrl(fileNameXlsx).getPath

  val basePath = "target"

  def randomInt: Int = Random.nextInt(10000)

  def fileExists(ruta: String): Boolean = new File(ruta).exists()

  def createCell(row: Row, column: Int, cellStyle: Option[CellStyle]): Cell = {
    val cell = row.createCell(column)
    cellStyle map (cs => {
      cell.setCellStyle(cs)
      cell
    }) getOrElse cell
  }

}
