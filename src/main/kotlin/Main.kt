import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

data class DataParse(
    val nSerie: String,
    val nCpf: String,
    val nValue: String,
    val nName: String
) {
    companion object {
        fun parse(element: Row) = DataParse(
            nSerie = element.getCell(1).toString(),
            nCpf = element.getCell(2).toString().replace("[.-]".toRegex(), ""),
            nValue = " ",
            nName = element.getCell(3).toString().uppercase()
        )
    }
}

fun readExcelFile(filePath: String) {
    val inputStream = FileInputStream(filePath)
    val xlWb = XSSFWorkbook(inputStream)
        .getSheetAt(0)
        .map {
            DataParse.parse(it.first().row)
        }

    println(xlWb)
}

fun main(args: Array<String>) {
    val filePath = "files/cartoes.xlsx"
    readExcelFile(filePath)
}