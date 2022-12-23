import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.OutputStream

data class DataParse(
    val nSerie: String,
    val nCpf: String,
    val nValue: String,
    val nName: String
) {
    companion object {
        fun parse(element: Row) = DataParse(
            nSerie = element
                .getCell(1)
                .toString(),
            nCpf = element
                .getCell(2)
                .toString()
                .replace("[.-]".toRegex(), ""),
            nValue = " ",
            nName = element
                .getCell(3)
                .toString()
                .let {
                    val (fName, sName) = it.split("\\s".toRegex()).toTypedArray()
                    "$fName $sName"
                }
                .trim()
                .uppercase()
        )
    }
}

fun parseExcelFile(filePath: String): List<DataParse> {
    try {
        val inputStream = FileInputStream(filePath)
        val xlWb = XSSFWorkbook(inputStream)
            .getSheetAt(0)
            .drop(1)
            .map {
                DataParse.parse(it.first().row)
            }
        return xlWb
    } catch (e: UnsupportedOperationException) {
        throw UnsupportedOperationException("Falha durante parse do arquivo de origem.")
    }
}

fun OutputStream.writeCsv(parsedRow: List<DataParse>) {
    try {
        val wr = bufferedWriter()
        wr.write("""Numero de Serie, CPF, Valor da Carga, Observacao""")
        wr.newLine()
        parsedRow.forEach {
            wr.write("${it.nSerie}, ${it.nCpf}, ${it.nValue}, ${it.nName}")
            wr.newLine()
        }
        wr.flush()
    } catch (e: UnsupportedOperationException) {
        throw UnsupportedOperationException("Falha durante a gravação do arquivo.")
    }
}

fun main(args: Array<String>) {
    val filePath = "files/cartoes.xlsx"
    val parsedRows = parseExcelFile(filePath)
    FileOutputStream("files/output.csv").apply { writeCsv(parsedRows) }
    println("Arquivo gerado com sucesso.")
}