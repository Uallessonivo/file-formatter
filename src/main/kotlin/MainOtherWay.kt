import au.com.bytecode.opencsv.CSVWriter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.FileWriter

data class ParseRow(
    val nSerie: String,
    val nCpf: String,
    val nName: String
)

fun main(args: Array<String>) {
    val inputStream = FileInputStream("files/cartoes.xlsx")
    val workbook = WorkbookFactory.create(inputStream)

    val sheet = workbook.getSheetAt(0)

    val csvFile = File("files/output_2.csv")
    val writer = CSVWriter(FileWriter(csvFile))

    writer.writeNext("""Numero de Serie, CPF, Valor da Carga, Observacao""")

    for (row in sheet.drop(1)) {
        val parsedRow = ParseRow(
            nSerie = row
                .getCell(1)
                .toString(),
            nCpf = row
                .getCell(2)
                .toString()
                .replace("[.-]".toRegex(), ""),
            nName = row
                .getCell(3)
                .toString()
                .let {
                    val (fName, sName) = it.split("\\s".toRegex()).toTypedArray()
                    "$fName $sName"
                }
                .trim()
                .uppercase()
        )
        writer.writeNext(parsedRow.nSerie, parsedRow.nCpf, "", parsedRow.nName)
    }
    inputStream.close()
    writer.close()
}
