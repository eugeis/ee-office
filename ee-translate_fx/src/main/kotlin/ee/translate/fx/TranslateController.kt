package ee.translate.fx

import ee.common.ext.withFileNameSuffix
import ee.pptx.PowerPoint
import ee.translate.TranslateServiceNoNeedTranslation
import ee.translate.TranslationServiceByGoogle
import ee.translate.TranslationServiceCsv
import ee.translate.pptx.translateTo
import javafx.application.Platform
import tornadofx.*
import java.io.File
import java.nio.file.Path
import java.nio.file.Paths

class TranslateController : Controller() {
    val dashboard: Dashboard by inject()
    val translations: MutableMap<String, String> = mutableMapOf()
    val translationServiceGoogle: TranslationServiceByGoogle = TranslationServiceByGoogle(config.string(LANGUAGE_FROM, "ru"),
            config.string(LANGUAGE_TO, "de"))
    val translationService: TranslationServiceCsv = TranslationServiceCsv(Paths.get(config.string(DICTIONARY, "dictionary.csv")),
            TranslateServiceNoNeedTranslation(translationServiceGoogle), translations)

    fun init() {
        showDashboard()
        translationService.load()
    }

    fun showDashboard() {
        if (FX.primaryStage.scene.root != dashboard.root) {
            FX.primaryStage.scene.root = dashboard.root
            FX.primaryStage.sizeToScene()
            FX.primaryStage.centerOnScreen()
            FX.primaryStage.title = dashboard.title
        }

        Platform.runLater {
            with(config) {
                if (containsKey(DICTIONARY)) {
                    dashboard.dictionary.text = string(DICTIONARY)
                }
                if (containsKey(LANGUAGE_FROM)) {
                    dashboard.languageFrom.text = string(LANGUAGE_FROM)
                }
                if (containsKey(LANGUAGE_TO)) {
                    dashboard.languageTo.text = string(LANGUAGE_TO)
                }
            }
        }
    }

    fun translate(files: List<File>, interactive: Boolean, translationUpdater: (String) -> Unit) {
        files.forEach { file ->
            val slideShow = PowerPoint.open(Paths.get(file.name).toFile())
            val target = file.name.withFileNameSuffix("_de")
            slideShow.translateTo(translationService, Paths.get(target).toFile())
        }

    }

    fun storeSettings() {
        with(config) {
            set(DICTIONARY to dashboard.dictionary.text)
            set(LANGUAGE_FROM, dashboard.languageFrom.text)
            set(LANGUAGE_TO, dashboard.languageTo.text)
            save()
        }
    }

    fun clearSettings() {
        with(config) {
            remove(DICTIONARY)
            remove(LANGUAGE_FROM)
            remove(LANGUAGE_TO)
            save()
        }
    }

    fun exit() {
        storeSettings()
        Platform.exit()
    }

    companion object {
        val TARGET_TO = "TARGET_TO"
        val DICTIONARY = "DICTIONARY"
        val LANGUAGE_FROM = "LANGUAGE_FROM"
        val LANGUAGE_TO = "LANGUAGE_TO"
    }

    fun changeDictionary(csvFile: Path) {
        translationService.csvFilePath = csvFile
        translationService.load()
    }

    fun changeFrom(languageFrom: String) {
        translationServiceGoogle.changeSource(languageFrom)
    }

    fun changeTo(languageTo: String) {
        translationServiceGoogle.changeSource(languageTo)
    }
}