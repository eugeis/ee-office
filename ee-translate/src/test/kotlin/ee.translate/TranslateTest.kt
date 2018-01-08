package ee.translate

fun main(args: Array<String>) {
    //val translate = Trans("ru", "de","AIzaSyCE0_thNhm-ryqW2pl6lEdxMoXpk9pYiV4")
    val translate = TranslationServiceByGoogle("ru", "de")
    val translated = translate.translate("привет")
    println(translated)
}