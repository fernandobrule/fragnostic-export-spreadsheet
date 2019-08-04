package com.fragnostic.export.excel

import java.util.Locale

import com.fragnostic.export.excel.support.{ BaseTest, SomeRow }
import org.apache.poi.ss.usermodel.Row

/**
 * Created by fernandobrule on 5/19/17.
 */
class ExportExcelTest extends BaseTest {

  describe("Export Excel Test") {

    it("Can Export Excel") {

      val list = List(
        SomeRow("Pepe", "+56 9 7979 7865"),
        SomeRow("John", "+56 9 3434 5678"))

      val sheetName: String = "pepe"
      val headers: Array[String] = Array("Name", "Telefono")
      val enables = (true, true, true)

      def newRow(locale: Locale, someRow: SomeRow, row: Row, enables: (Boolean, Boolean, Boolean)): Row = {
        row.createCell(0).setCellValue(someRow.name)
        row.createCell(1).setCellValue(someRow.tel)
        row
      }

      val wb = CakeService.exportExcelService.exportWb(list, sheetName, headers, enables, newRow) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val ruta = s"$basePath/workbook-$randomInt.xls"
      println(s"Can Export -> ruta: $ruta")
      fileExists(ruta) should be(false)

      save(wb, ruta)

      fileExists(ruta) should be(true)

    }

  }

}
