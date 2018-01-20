package ee.docx4j

import org.docx4j.dml.CTRegularTextRun
import org.docx4j.openpackaging.parts.DrawingML.DiagramDataPart
import org.docx4j.openpackaging.parts.DrawingML.DiagramDrawingPart
import org.slf4j.LoggerFactory


private val log = LoggerFactory.getLogger("PowerPointDocx4jTest")

fun main(args: Array<String>) {
    val ppt = PowerPoint.open("/Users/ee/Documents/Gemeinde/Seminare/Arts.pptx")

    ppt.parts.parts.forEach { name, part ->
        log.info("{},{}", name, part)
        when (part) {
            is DiagramDataPart -> {
                part.contents.ptLst.pt.forEach { pt ->
                    pt.t.p.forEach { p ->
                        val translationGroups = TextRunGroups(p)
                        translationGroups.groups.forEach { g ->
                            log.info("g: {} in {} in ", g, pt, part)
                        }
                    }
                    log.info("pt: {} in {}", pt, part)
                }
                log.info("DiagramDataPart: {},{}", name, part)
            }
            else -> log.info("Else: {},{}", name, part)
        }
    }


    val part = ppt.mainPresentationPart
    for (i in 0..part.slideCount) {
        val slide = part.getSlide(i)
        slide.slideLayoutPart
        //slide.contents.

    }
}