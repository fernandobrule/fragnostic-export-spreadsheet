package com.fragnostic.export.spreadsheet.support

import java.io.File

import org.apache.poi.ss.usermodel.{ Cell, CellStyle, Row }
import org.scalatest.{ FunSpec, Matchers }

import scala.util.Random
/**
 * Created by fernandobrule on 5/20/17.
 */
class BaseTest extends FunSpec with Matchers {

  val userUtcTime: String = ""
  val token: Option[String] = None

  val basePath = "target"

  def random = new Random()

  def randomInt = random.nextInt(10000)

  def fileExists(ruta: String) = new File(ruta).exists()

  def createCell(row: Row, column: Int, cellStyle: Option[CellStyle]): Cell = {
    val cell = row.createCell(column)
    cellStyle map (cs => {
      cell.setCellStyle(cs)
      cell
    }) getOrElse cell
  }

}
