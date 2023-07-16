import org.apache.commons.lang3.StringUtils
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.util.*
import javax.naming.Context
import javax.naming.NamingException
import javax.naming.directory.*
import javax.naming.ldap.Control
import javax.naming.ldap.PagedResultsControl


fun main(args: Array<String>) {
    val account = readFromXlsx("D:\\project\\compare\\user.xlsx")
    val enableAccount = account["enableAccount"]
    val disabledAccount = account["disabledAccount"]
    println("enableAccount(${enableAccount!!.size}): $enableAccount\ndisabledAccount(${disabledAccount!!.size}): $disabledAccount")

    val targetAccount = searchAccount()
    for (account in targetAccount){
        if (enableAccount.contains(account)
    }
    if (targetAccount. enableAccount)
}


fun readFromXlsx(filePath: String): MutableMap<String, MutableList<String>> {
    val file = File(filePath)
    val workbook = XSSFWorkbook(file)
    val sheet = workbook.getSheetAt(0) // Assuming you want to read the first sheet
    val enableAccount: MutableList<String> = mutableListOf()
    val disabledAccount: MutableList<String> = mutableListOf()

    for (rowIndex in 1..sheet.lastRowNum) { // Start from the second row
        val row = sheet.getRow(rowIndex)
        val account = when (row.getCell(0).cellType) {
            org.apache.poi.ss.usermodel.CellType.NUMERIC -> row.getCell(0).numericCellValue.toString()
            org.apache.poi.ss.usermodel.CellType.STRING -> row.getCell(0).stringCellValue
            org.apache.poi.ss.usermodel.CellType.BOOLEAN -> row.getCell(0).booleanCellValue.toString()
            else -> ""
        }
        val status = when (row.getCell(1).cellType) {
            org.apache.poi.ss.usermodel.CellType.NUMERIC -> row.getCell(1).numericCellValue.toString()
            org.apache.poi.ss.usermodel.CellType.STRING -> row.getCell(1).stringCellValue
            org.apache.poi.ss.usermodel.CellType.BOOLEAN -> row.getCell(1).booleanCellValue.toString()
            else -> ""
        }
        if (StringUtils.equalsIgnoreCase(status, "Enable")) {
            enableAccount.add(account)
        } else if (StringUtils.equalsIgnoreCase(status, "Disable")) {
            disabledAccount.add(account)
        }

    }
    workbook.close()
    return mutableMapOf (
        "enableAccount" to enableAccount,
        "disabledAccount" to disabledAccount
    )
}

fun searchAccount(): MutableMap<String, MutableList<String>> {
    return mutableMapOf (
        "enableAccount" to enableAccount,
        "disabledAccount" to disabledAccount
    )
    return mutableListOf("Wade", "Dave", "Seth", "Ivan", "Riley")
}