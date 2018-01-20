package ee.pptx

import ee.translate.translateFiles

fun main(args: Array<String>) {
    val sourceDir = "/Users/ee/Google Drive/Predigtreihe - David/0. Слайды Давид ЛФ"
    val targetDir = "/Users/ee/Documents/Bibelschule/Seminare/David/de"

    val fileTranslator = PptxFileTranslator()

    translateFiles(collectPowerPointFiles(sourceDir), targetDir, "$targetDir/dictionaryGlobal.xlsx",
        "$targetDir/dictionary.xlsx", "ru", "de", {}, true, { isColor(red = 255) }, fileTranslator)
}