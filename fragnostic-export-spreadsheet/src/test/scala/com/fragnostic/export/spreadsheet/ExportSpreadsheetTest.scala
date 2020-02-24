package com.fragnostic.export.spreadsheet

import java.util.Locale

import com.fragnostic.export.spreadsheet.support.{ BaseTest, SomeRow }
import org.apache.poi.ss.usermodel.{ Row, Workbook }
import com.fragnostic.export.spreadsheet.CakeServiceExportSpreadsheet.export

/**
 * Created by fernandobrule on 5/19/17.
 */
class ExportSpreadsheetTest extends BaseTest {

  describe("Export Spreadsheet Test") {

    it("Can Export Spreadsheet") {

      val locale: Locale = new Locale.Builder().setRegion("CL").setLanguage("es").build

      val list = List(
        SomeRow("Pepe", "+56 9 7979 7865"),
        SomeRow("John", "+56 9 3434 5678"))

      val sheetName: String = "pepe"
      val headers: Array[String] = Array("Name", "Telefono")

      def newRow(locale: Locale, someRow: SomeRow, row: Row): Row = {
        row.createCell(0).setCellValue(someRow.name)
        row.createCell(1).setCellValue(someRow.tel)
        row
      }

      val wb: Workbook = export.getWorkbook(locale, list, sheetName, headers, newRow) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val ruta = s"$basePath/workbook-$randomInt.xls"
      println(s"Can Export -> ruta: $ruta")
      fileExists(ruta) should be(false)

      export.save(wb, ruta)

      fileExists(ruta) should be(true)

    }

  }

}
