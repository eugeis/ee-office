package ee.docx4j

import ee.common.ext.collectFilesByExtension
import ee.translate.translateFiles

fun main(args: Array<String>) {
    val sourceDir = "/Users/ee/Documents/Gemeinde/Seminare/Arts.pptx"
    val targetDir = "/Users/ee/Documents/Bibelschule/Seminare/David/de"

    val fileTranslator = Docx4jPptxFileTranslator()

    translateFiles(collectFilesByExtension(sourceDir, ".pptx"), targetDir,
        "/Users/ee/Google Drive/Gemeinde/Übersetzung/dictionary_global.xlsx", "", "ru", "de",
        {}, true, { this.rPr.isColor("FF0000") }, fileTranslator)
}