package ee.docx4j

import ee.common.ext.collectFilesByExtension
import ee.translate.translateFiles

fun main(args: Array<String>) {
    val base = "/Users/ee/tmp"
    val sourceDir = "$base/Arts.pptx"
    val targetDir = "$base/de"

    val fileTranslator = Docx4jPptxFileTranslator()

    translateFiles(collectFilesByExtension(sourceDir, ".pptx"), targetDir,
        "/Users/ee/Google Drive/Gemeinde/Übersetzung/dictionary_global.xlsx", "", "ru", "de",
        {}, true, { this.rPr.isColor("FF0000") }, fileTranslator)
}