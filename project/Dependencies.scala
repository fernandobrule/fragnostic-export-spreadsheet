import sbt._
import sbt.Keys._

object Dependencies {

  private val scalatestVersion = "3.0.8"
  private val specs2Version = "4.6.0"
  private val poiVersion = "4.1.0"

  lazy val poi = "org.apache.poi" % "poi" % poiVersion
  lazy val poiOoxml = "org.apache.poi" % "poi-ooxml" % poiVersion
  lazy val slf4jApi = "org.slf4j" % "slf4j-api" % "1.7.26"
  lazy val scalatest = "org.scalatest" %% "scalatest" % scalatestVersion
  lazy val specs2 = Seq(
    "org.specs2" %% "specs2-core", 
    "org.specs2" %% "specs2-mock", 
    "org.specs2" %% "specs2-matcher-extra"
    ).map(_ % specs2Version)

}
