package com.fragnostic.spreadsheet

import com.fragnostic.spreadsheet.api.SpreadsheetServiceApi
import com.fragnostic.spreadsheet.impl.SpreadsheetServiceImpl

/**
 * Created by fernandobrule on 5/19/17.
 */
object CakeServiceSpreadsheet {

  lazy val spreadsheetPiece: SpreadsheetServiceApi = new SpreadsheetServiceImpl {}

  lazy val spreadsheet = spreadsheetPiece.spreadsheetService

}
