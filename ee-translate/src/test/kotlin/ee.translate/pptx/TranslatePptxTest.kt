package ee.translate.pptx

import org.apache.poi.sl.usermodel.PaintStyle

fun main(args: Array<String>) {
    translatePowerpoints("/Users/ee/Google Drive/Predigtreihe - David/0. Слайды Давид ЛФ",
            "/Users/ee/Documents/Bibelschule/Seminare", "ru", "de",
            "AIzaSyCE0_thNhm-ryqW2pl6lEdxMoXpk9pYiV4", {
        var ret = false
        if(fontColor is PaintStyle.SolidPaint) {
            val color = (fontColor as PaintStyle.SolidPaint).solidColor.color
            ret = color.red==255 && color.green==0 && color.blue==0
        }
        ret
    })
}