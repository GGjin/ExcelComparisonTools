package com.gg.exceltools.view

import javafx.application.Platform
import javafx.geometry.Insets
import javafx.geometry.Pos
import javafx.scene.Scene
import javafx.scene.control.Label
import javafx.scene.paint.Color
import javafx.scene.text.Font
import javafx.stage.Stage
import javafx.stage.StageStyle
import java.util.*

/*
 *@author unclezs.com
 *@date 2019.07.06 12:46
 */
object ToastUtil {
    private val stage = Stage()
    private val label = Label()
    /**
     * 指定时间消失
     * @param msg
     * @param time
     */
   //默认3秒
    @JvmOverloads
    fun toast(msg: String, time: Int = 3000) {
        label.text = msg
        val task: TimerTask = object : TimerTask() {
            override fun run() {
                Platform.runLater { stage.close() }
            }
        }
        init(msg)
        val timer = Timer()
        timer.schedule(task, time.toLong())
        stage.show()
    }

    //设置消息
    private fun init(msg: String) {
        val label = Label(msg) //默认信息
        label.style = "-fx-background: rgba(56,56,56,0.7);-fx-border-radius: 25;-fx-background-radius: 25" //label透明,圆角
        label.textFill = Color.rgb(225, 255, 226) //消息字体颜色
        label.prefHeight = 50.0
        label.padding = Insets(15.0)
        label.alignment = Pos.CENTER //居中
        label.font = Font(20.0) //字体大小
        val scene = Scene(label)
        scene.fill = null //场景透明
        stage.scene = scene
    }

    init {
        stage.initStyle(StageStyle.TRANSPARENT) //舞台透明
    }
}