package com.zorenkonte.exceltokml

import android.content.Context
import android.net.Uri
import androidx.core.content.FileProvider
import androidx.lifecycle.LiveData
import androidx.lifecycle.MutableLiveData
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.*
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

class ExcelViewModel : ViewModel() {
    private val _progress = MutableLiveData<Int>()
    val progress: LiveData<Int> get() = _progress

    private val _data = MutableLiveData<List<List<String>>>()
    val data: LiveData<List<List<String>>> get() = _data

    fun readExcelFile(context: Context, uri: Uri) {
        viewModelScope.launch(Dispatchers.IO) {
            val inputStream: InputStream? = context.contentResolver.openInputStream(uri)
            val formatter = DataFormatter() // Format cell values (including empty cells)
            val rows = mutableListOf<List<String>>()

            inputStream.use { stream ->
                val bufferedStream = BufferedInputStream(stream)
                val workbook = WorkbookFactory.create(bufferedStream)
                val sheet = workbook.getSheetAt(0)
                val totalRows = sheet.physicalNumberOfRows - 1

                // Iterate over each row
                sheet.drop(1).forEachIndexed { rowIndex, row ->
                    val rowData = mutableListOf<String>()
                    // Ensure that we process all cells, including empty ones
                    for (cellIndex in 0 until row.lastCellNum) {
                        val cell = row.getCell(cellIndex) // Get cell at this index
                        val cellValue = formatter.formatCellValue(cell) // Format the cell value
                        rowData.add(cellValue) // Add the value (even if it's empty)
                    }
                    rows.add(rowData)
                    updateProgress(rowIndex, totalRows)
                }
            }

            _data.postValue(rows)
        }
    }


    private fun updateProgress(rowIndex: Int, totalRows: Int) {
        _progress.postValue((rowIndex + 1) * 100 / totalRows)
    }

    fun resetProgress() {
        _progress.value = 0
    }

    fun convertToKML(data: List<List<String>>, context: Context): Uri? {
        val kmlFile = writeKMLToStream(data, context) // Write KML progressively
        return kmlFile?.let { compressToKMZ(it, context) } // Compress to KMZ after writing
    }

    private fun writeKMLToStream(data: List<List<String>>, context: Context): File? {
        return try {
            val file = File(context.cacheDir, "output.kml")
            BufferedWriter(FileWriter(file)).use { writer ->
                writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                writer.write("<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n")
                writer.write("<Document>\n")

                data.drop(1).forEach { row ->
                    writer.write(createPlaceMark(row)) // Write each placemark in chunks
                }

                writer.write("</Document>\n")
                writer.write("</kml>\n")
            }
            file
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    private fun compressToKMZ(kmlFile: File, context: Context): Uri? {
        return try {
            val kmzFile = File(context.cacheDir, "output.kmz")
            ZipOutputStream(FileOutputStream(kmzFile)).use { zos ->
                zos.putNextEntry(ZipEntry(kmlFile.name)) // Add KML to KMZ
                FileInputStream(kmlFile).use { fis ->
                    fis.copyTo(zos) // Stream KML content to KMZ
                }
                zos.closeEntry() // Finish entry
            }
            FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", kmzFile)
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    private fun createPlaceMark(row: List<String>): String {
        return """
            <Placemark>
                <name>${row[0]}</name>
                <description><![CDATA[
                    <p>CAN: ${row[0]}<br/>
                       Meter Code: ${row[1]}<br/>
                       MRU: ${row[2]}<br/>
                       BA: ${row[3]}<br/>
                       Name: ${row[4]}<br/>
                       Address: ${row[5]}<br/>
                       DMA: ${row[6]}<br/>
                       DMZ: ${row[7]}<br/>
                       Latitude: ${row[9]}<br/>
                       Longitude: ${row[10]}<br/>
                       Coordinates Source: ${row[11]}</p>
                ]]></description>
                <Point><coordinates>${row[10]},${row[9]}</coordinates></Point>
            </Placemark>
        """.trimIndent()
    }
}
