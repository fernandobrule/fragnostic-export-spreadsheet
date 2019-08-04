package com.fragnostic.export.excel

import com.fragnostic.export.excel.impl.ExportExcelServiceImpl

/**
 * Created by fernandobrule on 5/19/17.
 */
object CakeService {

  lazy val exportExcelServicePiece = new ExportExcelServiceImpl {
  }

  val exportExcelService = exportExcelServicePiece.exportExcelService

}
