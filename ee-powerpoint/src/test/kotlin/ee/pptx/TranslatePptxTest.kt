package ee.pptx

import ee.common.ext.collectFilesByExtension
import ee.translate.translateFiles

fun main(args: Array<String>) {
    val sourceDir = "/Users/ee/Documents/Gemeinde/Seminare/Arts.pptx"
    val targetDir = "/Users/ee/Documents/Bibelschule/Seminare/David/de"

    val fileTranslator = PptxFileTranslator()

    translateFiles(collectFilesByExtension(sourceDir, ".pptx"), targetDir, "/Users/ee/Google Drive/Gemeinde/Übersetzung/dictionary_global.xlsx",
        "$targetDir/dictionary.xlsx", "ru", "de", {}, true, { isColor(red = 255) }, fileTranslator)
}