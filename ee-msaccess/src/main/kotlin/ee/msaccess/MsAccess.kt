package ee.msaccess

import com.healthmarketscience.jackcess.Database
import com.healthmarketscience.jackcess.DatabaseBuilder
import java.nio.file.Paths

class MsAccess {
    companion object {
        @JvmStatic
        fun open(fileName: String): Database {
            return DatabaseBuilder.open(Paths.get(fileName).toFile())
        }
    }
}