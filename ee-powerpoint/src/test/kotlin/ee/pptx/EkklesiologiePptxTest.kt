package ee.pptx

import com.fasterxml.jackson.module.kotlin.readValue
import ee.common.ext.exists
import ee.common.ext.isWindows
import ee.slides.Presentation
import ee.slides.extractPicturesTo
import ee.slides.html.toCss
import ee.slides.html.toCssNamesAll
import ee.slides.html.toHtml
import ee.slides.html.toReveal
import ee.slides.mapper
import java.nio.file.Path
import java.nio.file.Paths


fun main(args: Array<String>) {
    val path = Paths.get(if (isWindows)
        "G:/Ekklesiologie/Seminar"
    else
        "/home/z000ru5y/data/rclone/eo/Bibelschule/Stephanus/Ekklesiologie/Seminar")

    val target = path.resolve("reveal")

    val presentation: Presentation = loadOrParseToJson(path, target)

    generate(presentation, target)
}

private fun generate(presentation: Presentation, target: Path) {
    target.resolve("${presentation.name}-slides2.html").toFile().writeText(presentation.toReveal().toString())
    target.resolve("${presentation.name}-script2.html").toFile().writeText(presentation.toHtml().toString())
    target.resolve("powerpoint.names.css").toFile().writeText(presentation.toCssNamesAll().toCss().toString())
}

private fun loadOrParseToJson(pptxPath: Path, target: Path): Presentation {
    val jsonFile = pptxPath.resolve("slides.json")
    val mapper = mapper()
    val presentation: Presentation
    if (!jsonFile.exists()) {
        presentation = PowerPoint.parseFilesAsTopics(pptxPath, "Ekklesiologie")
        presentation.extractPicturesTo(target)
        mapper.writeValue(jsonFile.toFile(), presentation)
    } else {
        presentation = mapper.readValue(jsonFile.toFile())
    }
    return presentation
}
