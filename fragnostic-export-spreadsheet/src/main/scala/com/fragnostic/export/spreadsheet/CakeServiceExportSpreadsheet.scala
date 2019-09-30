package com.fragnostic.export.spreadsheet

import com.fragnostic.export.spreadsheet.impl.ExportSpreadsheetServiceImpl

/**
 * Created by fernandobrule on 5/19/17.
 */
object CakeServiceExportSpreadsheet {

  lazy val exportPiece = new ExportSpreadsheetServiceImpl {}

  val export = exportPiece.exportSpreadsheetService

}
