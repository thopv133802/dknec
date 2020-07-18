package com.thopham.projects.dknec.demo_v2

import com.opencsv.CSVReader
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.core.io.ClassPathResource
import java.io.FileReader
import java.util.concurrent.ThreadLocalRandom

//fun main(args: Array<String>) {
//    val tags = (nau() + men())
//            .filter { tag -> !tag.contains("Glycol") && !tag.contains(" 1") && !tag.contains(" 2") }
//    val tags = nau() + men()
//    print(tags.size)
//    for (tag in tags) {
//        println("""create index z_tag_${tag}_created_index on z_tag_$tag (created desc);""")
//    }
//    generate_create_tables(tags)
//}

private fun generate_create_tables(tags: List<String>) {
    val query = tags.map { tag ->
        """
create table z_tag_$tag (tag_value real, created datetime);
    """.trim()
    }.joinToString("\n")
    println(query)
}

private fun generate_vbcode(tags: List<String>) {
    val vbcode = tags.map { tag ->
        """
Dim tag_$tag
tag_$tag = HMIRuntime.Tags("$tag").Read
strSQL = "INSERT INTO z_tag_$tag (tag_value, created) values(" & tag_$tag & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_$tag = Nothing
""".trim()
    }
            .joinToString("\n")
    print(vbcode)
}

private fun nau(): List<String> {
    val excelFilePath = "Export_nau.xlsx"
    val inputStream = ClassPathResource(excelFilePath).inputStream
    val workbook = XSSFWorkbook(inputStream)
    val tagsSheet = workbook.getSheetAt(2)
    val rows = tagsSheet.iterator()

    val tags = mutableListOf<String>()

    while (rows.hasNext()) {
        val nextRow = rows.next()
        val cells = nextRow.cellIterator()
        val tagName = cells.next().stringCellValue
        tags.add(tagName)
    }

    val filteredTags = tags.filter { tag ->
        val tagLowerCase = tag.toLowerCase()
        !tagLowerCase.contains("_sp") && !tagLowerCase.contains("_dfm")
    }.drop(3)

    return filteredTags
}

private fun men(): List<String> {
    val excelFilePath = "Export_men.xlsx"
    val inputStream = ClassPathResource(excelFilePath).inputStream
    val workbook = XSSFWorkbook(inputStream)
    val tagsSheet = workbook.getSheetAt(2)
    val rows = tagsSheet.iterator()

    val tags = mutableListOf<String>()

    while (rows.hasNext()) {
        val nextRow = rows.next()
        val cells = nextRow.cellIterator()
        val tagName = cells.next().stringCellValue
        tags.add(tagName)
    }

    val filteredTags = tags.filter { tag ->
        val tagLowercase = tag.toLowerCase()
        !tagLowercase.contains("_sp") && !tagLowercase.contains("_dfm")
    }.drop(3)

    val query = filteredTags.joinToString("\n") { tag ->
        "create table tag_$tag (tag_name varchar(255), tag_value real, created varchar(255));"
    }
    return filteredTags
}
private fun chiet(): List<String>{
    val tags = mutableListOf<String>()
    val reader = CSVReader(FileReader(ClassPathResource("chietchai_vex.csv").file))
    while(true) {
        val rows = reader.readNext()
        if(rows.isNullOrEmpty())
            break
        tags.add(rows[0])
    }
    reader.close()
    return tags.filter { tag -> !tag.contains("@") }.drop(3)
}