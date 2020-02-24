package com.fragnostic.export.spreadsheet

import com.fragnostic.export.spreadsheet.support.BaseTest
import org.apache.poi.ss.usermodel.Workbook
import CakeServiceExportSpreadsheet.export

class OrthogonalOnXlsTest extends BaseTest {

  describe("Orthogonal On Xls Test") {

    it("Can Be Orthogonal XLS") {

      val pathXls: String = "/Users/fernandobrule/Documents/alphanum.xls"

      val wb: Workbook = export.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue should be("asda fdg fd ertert yuyut ty")

      //https://github.com/pathikrit/better-files/
      import better.files.{ File => BFFile }
      val bytes: Array[Byte] = BFFile(pathXls).byteArray

      val bytesLength: Int = bytes.length
      println(s"xls bytesLength.length:$bytesLength")

      val path = s"$pathXls---ortho.xls"
      export.save(bytes, path)

      //getBytes(getWorkbook(path)).length should be(bytesLength)
      BFFile(path).byteArray.length should be(bytesLength)

    }

  }

}