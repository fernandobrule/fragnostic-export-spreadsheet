import sbt._
import sbt.Keys._

object Dependencies {

  private lazy val poiVersion = "4.1.0"

  lazy val logbackClassic = "ch.qos.logback" % "logback-classic" % "1.2.3" % "runtime"
  lazy val poi = "org.apache.poi" % "poi" % poiVersion
  lazy val poiOoxml = "org.apache.poi" % "poi-ooxml" % poiVersion
  lazy val scalatest = "org.scalatest" %% "scalatest" % "3.0.8" % "test"
  lazy val slf4jApi = "org.slf4j" % "slf4j-api" % "1.7.26"
  lazy val betterFiles = "com.github.pathikrit" %% "better-files" % "3.8.0" % "test"

}
