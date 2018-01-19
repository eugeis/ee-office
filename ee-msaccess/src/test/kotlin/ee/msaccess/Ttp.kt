package ee.msaccess

import org.slf4j.LoggerFactory

fun main(args: Array<String>) {
    val log = LoggerFactory.getLogger("ttp")
    val ttp = MsAccess.open("D:\\Vicos\\MetroGalaxy\\TTP\\db.mdb")
    ttp.tableNames.forEach {
        val table = ttp.getTable(it)
        log.info("{}", table.name)
        table.forEach {
            log.info("{}", it)
        }
    }
    ttp.close()

}