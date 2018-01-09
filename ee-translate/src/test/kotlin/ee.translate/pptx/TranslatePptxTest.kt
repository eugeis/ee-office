package ee.translate.pptx

fun main(args: Array<String>) {
    val sourceDir = "/Users/ee/Google Drive/Predigtreihe - David/0. Слайды Давид ЛФ"
    val targetDir = "/Users/ee/Documents/Bibelschule/Seminare/David/de"

    translatePowerPoints(sourceDir, targetDir,
            "$targetDir/dictionaryGlobal.xlsx", "$targetDir/dictionary.xlsx",
            "ru", "de", {}, { isColor(red = 255.0) })
}