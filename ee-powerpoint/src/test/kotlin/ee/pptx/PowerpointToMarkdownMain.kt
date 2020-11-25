package ee.pptx

import com.fasterxml.jackson.module.kotlin.readValue
import ee.common.ext.exists
import ee.slides.Presentation
import ee.slides.extractPicturesTo
import ee.slides.html.toCss
import ee.slides.html.toCssNamesAll
import ee.slides.html.toHtml
import ee.slides.html.toReveal
import ee.slides.mapper
import ee.slides.markdown.SlideToMarkdown
import ee.slides.markdown.SlideToVSCodeReveal
import java.nio.file.Path
import java.nio.file.Paths

fun main() {
    val path = "/home/z000ru5y/data/cloud/Seminare/Faulheit"
    val name = "Faulheit"

    generate(path, name)
}

private fun generate(path: String, name: String) {
    val source = Paths.get(path)
    val pres = loadOrParseToJson(source, name)
    generate(pres, source)
}

private fun generate(presentation: Presentation, target: Path) {
    target.resolve("${presentation.name}-vscode-reveal.md").toFile().writeText(
            SlideToVSCodeReveal(presentation).generate().toString())
    target.resolve("${presentation.name}-slides.md").toFile().writeText(
            SlideToMarkdown(presentation).generate().toString())
    target.resolve("${presentation.name}-slides.html").toFile().writeText(presentation.toReveal().toString())
    target.resolve("${presentation.name}-script.html").toFile().writeText(presentation.toHtml().toString())
    target.resolve("powerpoint.names.css").toFile().writeText(presentation.toCssNamesAll().toCss().toString())
}

private fun loadOrParseToJson(pptxFolder: Path, name: String): Presentation {
    val jsonFile = pptxFolder.resolve("${name}_slides.json")
    val mapper = mapper()
    val presentation: Presentation
    if (!jsonFile.exists()) {
        presentation = PowerPoint.parseFilesAsTopics(pptxFolder, name)
        presentation.extractPicturesTo(pptxFolder.resolve("images"))
        mapper.writeValue(jsonFile.toFile(), presentation)
    } else {
        presentation = mapper.readValue(jsonFile.toFile())
        //presentation = presentation.aggregate()
        //mapper.writeValue(pptxPath.resolve("${name}_slides_aggregated.json").toFile(), presentation)
    }
    return presentation
}
