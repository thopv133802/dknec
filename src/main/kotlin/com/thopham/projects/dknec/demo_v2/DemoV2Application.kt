package com.thopham.projects.dknec.demo_v2

import org.springframework.boot.SpringApplication
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.builder.SpringApplicationBuilder
import org.springframework.cloud.openfeign.EnableFeignClients
import org.springframework.core.io.ClassPathResource
import java.awt.*
import java.awt.event.WindowAdapter
import java.awt.event.WindowEvent
import java.util.*
import javax.swing.JFrame
import javax.swing.WindowConstants
import kotlin.system.exitProcess


@SpringBootApplication
@EnableFeignClients
class DemoV2Application: JFrame() {
    companion object {
        @JvmStatic
        fun main(args: Array<String>) {
            val context = SpringApplicationBuilder(DemoV2Application::class.java).headless(false).run(*args)
            EventQueue.invokeLater {
                val uiApp = context.getBean(DemoV2Application::class.java)
                uiApp.isVisible = true
            }
        }
    }
    lateinit var trayIcon: TrayIcon
    init {
        setupUI()
        setupTrayIcon()
    }
    private fun setupTrayIcon() {
        val systemTrayIsNotSupported = !SystemTray.isSupported()
        if (systemTrayIsNotSupported) {
            println("SystemTray is not supported")
            return
        }
        val popupMenu = PopupMenu()
        val greetingItem = MenuItem("DKNec")
        popupMenu.add(greetingItem)
        val trayIconImage = Toolkit.getDefaultToolkit().createImage(ClassPathResource("logo.png").url)
        trayIcon = TrayIcon(trayIconImage, "DKNEC", popupMenu)
        addWindowListener(object: WindowAdapter() {
            override fun windowClosing(event: WindowEvent) {
                SystemTray.getSystemTray().remove(trayIcon)
                SystemTray.getSystemTray().add(trayIcon)
            }
        })
    }


    private fun setupUI() {
        title = "Hello World"
        setSize(450, 150)
        setLocationRelativeTo(null)
        defaultCloseOperation = WindowConstants.DO_NOTHING_ON_CLOSE
    }
}