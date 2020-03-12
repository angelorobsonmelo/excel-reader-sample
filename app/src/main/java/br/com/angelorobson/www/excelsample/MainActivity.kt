package br.com.angelorobson.www.excelsample

import android.os.Bundle
import android.util.Log
import androidx.appcompat.app.AppCompatActivity
import androidx.fragment.app.FragmentActivity
import kotlinx.android.synthetic.main.activity_main.*
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.Row
import java.io.InputStream


class MainActivity : AppCompatActivity() {

    var TAG = "main"


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        readExcelFileFromAssets()
    }

    fun readExcelFileFromAssets() {
        try {
            val myInput: InputStream
            // initialize asset manager
            val assetManager = assets
            //  open excel sheet
            myInput = assetManager.open("myexcelsheet.xls")
            // Create a POI File System object
            val myFileSystem = POIFSFileSystem(myInput)
            // Create a workbook using the File System
            val myWorkBook = HSSFWorkbook(myFileSystem)
            // Get the first sheet from workbook
            val mySheet = myWorkBook.getSheetAt(0)
            // We now need something to iterate through the cells.
            val rowIter: Iterator<Row> = mySheet.rowIterator()
            var rowno = 0
            textView.append("\n")
            while (rowIter.hasNext()) {
                Log.e(TAG, " row no $rowno")
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter = myRow.cellIterator()
                    var colno = 0
                    var sno = ""
                    var date = ""
                    var det = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        if (colno == 0) {
                            sno = myCell.toString()
                        } else if (colno == 1) {
                            date = myCell.toString()
                        } else if (colno == 2) {
                            det = myCell.toString()
                        }
                        colno++
                        Log.e(TAG, " Index :" + myCell.columnIndex + " -- " + myCell.toString())
                    }
                    if (sno.isNotEmpty() && date.isNotEmpty() && det.isNotEmpty()) {
                        textView.append("$sno -- $date  -- $det\n")
                    }

                }
                rowno++
            }
        } catch (e: Exception) {
            Log.e(TAG, "error $e")
        }
    }

}
