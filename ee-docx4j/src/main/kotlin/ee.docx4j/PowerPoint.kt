package ee.docx4j

import ee.common.ext.addReturn
import ee.translate.TranslationService
import org.docx4j.dml.*
import org.docx4j.openpackaging.packages.PresentationMLPackage
import org.slf4j.LoggerFactory
import java.io.File
import java.nio.file.Paths

private val log = LoggerFactory.getLogger("PowerPointDocx4j")


class PowerPoint {
    companion object {
        @JvmStatic
        fun open(fileName: String): PresentationMLPackage {
            return open(Paths.get(fileName).toFile())
        }

        @JvmStatic
        fun open(file: File): PresentationMLPackage {
            return PresentationMLPackage.load(file)
        }
    }
}


class TextGroups(val paragraph: CTTextParagraph, val items: MutableList<TextGroup> = mutableListOf(),
    private var text: String? = null) {
    init {
        if (paragraph.egTextRun.isNotEmpty()) {
            var currentTextRunGroup = items.addReturn(TextGroup(paragraph))

            paragraph.egTextRun.forEach {
                if (!currentTextRunGroup.addIfSimilar(it)) {
                    currentTextRunGroup = items.addReturn(TextGroup(paragraph))
                    currentTextRunGroup.addIfSimilar(it)
                }
            }
        }
    }

    fun text(): String {
        if (text == null) {
            text = items.joinToString("") { it.text() }
        }
        return text!!
    }
}

class TextPart(val paragraph: CTTextParagraph, private var text: String? = null) {
    var first: CTRegularTextRun? = null
    val otherTextRuns: MutableList<CTRegularTextRun> = mutableListOf()

    fun add(textRun: CTRegularTextRun) {
        if (first == null) {
            first = textRun
        } else {
            otherTextRuns.add(textRun)
        }
    }

    fun text(): String {
        if (text == null) {
            val out = StringBuilder()
            out.append(first?.t)
            for (r in otherTextRuns) {
                out.append(r.t)
            }
            text = out.toString()
        }
        return text!!
    }

    fun changeText(text: String) {
        val currentBaseTextRun = first
        if (currentBaseTextRun == null) {
            log.info("can't change text, because first is null")
            return
        }

        currentBaseTextRun.t = text
        currentBaseTextRun.rPr.lang = null

        //to check, maybe line break textRun is not needed because calculated bei PowerPoint
        otherTextRuns.forEach { paragraph.remove(it) }
        paragraph.egTextRun.removeAll(otherTextRuns)

        otherTextRuns.clear()
    }

    fun removeAll() {
        val currentBaseTextRun = first
        if (currentBaseTextRun != null) {
            paragraph.remove(currentBaseTextRun)
            paragraph.egTextRun.remove(currentBaseTextRun)

            otherTextRuns.forEach { paragraph.remove(it) }
            paragraph.egTextRun.removeAll(otherTextRuns)

            otherTextRuns.clear()
        }
    }
}

class TextGroup(val paragraph: CTTextParagraph, private var text: String? = null) {
    private var added = false
    private var currentTextPart: TextPart = TextPart(paragraph)
    val textParts: MutableList<TextPart> = mutableListOf()
    private var baseTextRun: CTRegularTextRun? = null

    fun addIfSimilar(textRun: Any): Boolean {
        var ret = true
        if (textRun.isLineBreak()) {
            if (added) {
                currentTextPart = TextPart(paragraph)
                added = false
            }
        } else if (textRun is CTRegularTextRun) {
            if (baseTextRun == null) {
                baseTextRun = textRun
                currentTextPart.add(textRun)
            } else if (baseTextRun!!.rPr.isSimilar(textRun.rPr)) {
                currentTextPart.add(textRun)
            } else {
                ret = false
            }
            if (ret && !added) {
                textParts.add(currentTextPart)
                added = true
            }
        }
        return ret
    }

    fun text(): String {
        if (text == null) {
            text = textParts.joinToString(TranslationService.NEW_LINE) { it.text() }
        }
        return text!!
    }

    fun changeText(text: String) {
        val translationParts = text.split(TranslationService.NEW_LINE)
        textParts.forEachIndexed { i, textPart ->
            if (i < translationParts.size) {
                textPart.changeText(translationParts[i])
            } else {
                textPart.removeAll()
            }
        }
    }

    fun removeAllFromParagraph(isToRemove: CTRegularTextRun.() -> Boolean): Boolean {
        val currentBaseTextRun = baseTextRun
        val ret = currentBaseTextRun != null && currentBaseTextRun.isToRemove()
        if (ret) {
            log.info("Remove text '{}' from paragraph '{}'", text(), paragraph)
            textParts.forEach { it.removeAll() }
        }
        return ret
    }
}

private fun CTTextParagraph.remove(textRun: Any) {
    val items = egTextRun
    for (index in 0..items.size - 1) {
        val item = items[index]
        if (textRun == item) {
            egTextRun.remove(textRun)
            break
        }
    }
}

fun Any.isLineBreak(): Boolean {
    return this is CTTextLineBreak
}

fun CTTextCharacterProperties.isSimilar(other: CTTextCharacterProperties): Boolean {
    var ret = true
    //ret = ret && this.ln == other.ln
    ret = ret && this.noFill == other.noFill
    ret = ret && this.solidFill.isSimilar(other.solidFill)
    ret = ret && this.gradFill == other.gradFill
    ret = ret && this.blipFill == other.blipFill
    ret = ret && this.pattFill == other.pattFill
    ret = ret && this.grpFill == other.grpFill
    ret = ret && this.effectDag == other.effectDag
    ret = ret && this.highlight == other.highlight
    ret = ret && this.uLn == other.uLn
    ret = ret && this.uFill == other.uFill
    ret = ret && this.sym == other.sym
    ret = ret && this.hlinkClick == other.hlinkClick
    ret = ret && this.hlinkMouseOver == other.hlinkMouseOver
    ret = ret && this.extLst == other.extLst
    ret = ret && this.isKumimoji.isSimilar(other.isKumimoji)
    ret = ret && this.altLang == other.altLang
    ret = ret && this.sz == other.sz
    ret = ret && this.isB == other.isB
    ret = ret && this.isI.isSimilar(other.isI)
    ret = ret && this.u.isSimilar(other.u)
    ret = ret && this.strike.isSimilar(other.strike)
    //ret = ret && this.kern == other.kern
    ret = ret && this.cap.isSimilar(other.cap)
    ret = ret && this.spc.isSimilar(other.spc)
    ret = ret && this.isNormalizeH.isSimilar(other.isNormalizeH)
    ret = ret && this.baseline.isSimilar(other.baseline)
    ret = ret && this.isNoProof.isSimilar(other.isNoProof)
    //ret = ret && this.isDirty.isSimilar(other.isDirty)
    ret = ret && this.smtId == other.smtId
    ret = ret && this.bmk == other.bmk
    //ret = ret && this.isErr == other.isErr
    //ret = ret && this.isSmtClean == other.isSmtClean
    //ret = ret && this.lang == other.lang
    //ret = ret && this.latin == other.latin
    //ret = ret && this.effectLst == other.effectLst
    //ret = ret && this.uFillTx == other.uFillTx
    //ret = ret && this.uLnTx == other.uLnTx
    //ret = ret && this.ea == other.ea
    //ret = ret && this.cs == other.cs

    return ret
}

fun CTSolidColorFillProperties?.isSimilar(other: CTSolidColorFillProperties?): Boolean {
    val ret = this == other || this?.srgbClr?.`val` == other?.srgbClr?.`val`
    return ret
}

fun Boolean?.isSimilar(other: Boolean?): Boolean {
    val ret = this == other || (this == null && other == false) || (other == null && this == false)
    return ret
}

fun STTextStrikeType?.isSimilar(other: STTextStrikeType?): Boolean {
    val ret = this == other || (this == null && other == STTextStrikeType.NO_STRIKE) ||
            (other == null && this == STTextStrikeType.NO_STRIKE)
    return ret
}

fun STTextCapsType?.isSimilar(other: STTextCapsType?): Boolean {
    val ret = this == other || (this == null && other == STTextCapsType.NONE) ||
            (other == null && this == STTextCapsType.NONE)
    return ret
}

fun STTextUnderlineType?.isSimilar(other: STTextUnderlineType?): Boolean {
    val ret = this == other || (this == null && other == STTextUnderlineType.NONE) ||
            (other == null && this == STTextUnderlineType.NONE)
    return ret
}

fun Int?.isSimilar(other: Int?): Boolean {
    val ret = this == other || (this == null && other == 0) ||
            (other == null && this == 0)
    return ret
}