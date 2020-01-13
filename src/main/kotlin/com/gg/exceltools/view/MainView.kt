package com.gg.exceltools.view

import com.gg.exceltools.app.Styles
import tornadofx.*

class MainView : View("Hello TornadoFX") {


    override val root = hbox {
        label(title) {
            addClass(Styles.heading)
        }
    }
}