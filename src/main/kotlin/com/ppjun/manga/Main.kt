package com.ppjun.manga

import com.alibaba.excel.EasyExcel
import com.alibaba.excel.write.metadata.fill.FillWrapper
import org.jsoup.Jsoup
import org.openqa.selenium.chrome.ChromeDriver

/**
 *  @author ppjun
 *  @date 2021/07/13.
 *  抓取动漫网站作品的名字写进excel
 */

class Main {
    private val chromeDriver = ChromeDriver()
    private val mangaCabinetList: ArrayList<Data> = arrayListOf()
    private val mangaHomeList: ArrayList<Data> = arrayListOf()

    fun fetchMangaHome() {
        repeat(904) {
            val page = it + 1
            chromeDriver.get("http://manhua.dmzj.com/tags/category_search/0-0-2304-all-0-1-0-${page}.shtml#category_nav_anchor")
            val doc = Jsoup.parse(chromeDriver.pageSource)
            val elements = doc.getElementsByClass("tcaricature_block tcaricature_block2")
            elements.forEach {
                val ul = it.select("ul")
                ul.forEach {
                    val name = it.selectFirst("li a")?.text() ?: ""
                    mangaHomeList.add(Data(name))
                }
            }
        }
    }

    fun writeToExcel() {
        val templeFileName = "manga.xlsx"
        val fileName = "manga-${System.currentTimeMillis()}.xlsx"
        val excelWriter = EasyExcel.write(fileName).withTemplate(templeFileName).build()
        val writeSheet = EasyExcel.writerSheet().build()
        excelWriter.fill(FillWrapper("data1", mangaHomeList), writeSheet)
        excelWriter.fill(FillWrapper("data2", mangaCabinetList), writeSheet)
        excelWriter.finish()
    }


    fun fetchMangaCabinet() {
        repeat(817) {
            val page = it + 1
            chromeDriver.get("https://www.manhuagui.com/list/japan/index_p${page}.html")
            val doc = Jsoup.parse(chromeDriver.pageSource)
            val elements = doc.getElementsByClass("book-list")
            elements.forEach {
                val ul = it.select("ul li")
                ul.forEach {
                    val name = it.selectFirst("p a")?.text() ?: ""
                    mangaCabinetList.add(Data(name))
                }
            }
        }
    }
}

fun main() {
    println("start to fetch")
    val main = Main()
    main.fetchMangaHome()
    main.fetchMangaCabinet()
    main.writeToExcel()
    println("fetch end")
}