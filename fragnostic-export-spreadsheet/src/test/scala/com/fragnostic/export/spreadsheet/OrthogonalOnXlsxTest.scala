package com.fragnostic.export.spreadsheet

import com.fragnostic.export.spreadsheet.CakeServiceExportSpreadsheet.export
import com.fragnostic.export.spreadsheet.support.BaseTest
import org.apache.poi.ss.usermodel.Workbook

class OrthogonalOnXlsxTest extends BaseTest {

  describe("Orthogonal On Xlsx Test") {

    it("Can Be Orthogonal On XLSX") {

      val pathXlsx: String = "/Users/fernandobrule/Documents/alphanum.xlsx"

      val wb: Workbook = export.getWorkbook(pathXlsx) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue should be("214312 345 456 5765756 qweq")

      // https://github.com/pathikrit/better-files/
      import better.files.{ File => BFFile }
      val bytes: Array[Byte] = BFFile(pathXlsx).byteArray

      val bytesLength: Int = bytes.length
      println(s"xls bytes.length:$bytesLength")

      val path = s"$pathXlsx---ortho.xlsx"
      export.save(bytes, path)

      //export.getBytes(getWorkbook(path)).length should be(bytesLength)
      BFFile(path).byteArray.length should be(bytesLength)
    }

  }

}