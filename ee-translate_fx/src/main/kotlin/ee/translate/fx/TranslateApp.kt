package ee.translate.fx

import javafx.event.EventHandler
import javafx.stage.Stage
import tornadofx.*

open class TranslateApp : App() {
    override val primaryView = Dashboard::class
    private val translateController: TranslateController by inject()

    override fun start(stage: Stage) {
        stage.onCloseRequest = EventHandler { translateController.exit() }
        importStylesheet(Styles::class)
        super.start(stage)
        translateController.init()
    }
}