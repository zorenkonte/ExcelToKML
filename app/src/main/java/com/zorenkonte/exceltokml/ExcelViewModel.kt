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
import java.io.BufferedInputStream
import java.io.File
import java.io.FileOutputStream
import java.io.InputStream

class ExcelViewModel : ViewModel() {
    private val _progress = MutableLiveData<Int>()
    val progress: LiveData<Int> get() = _progress

    private val _data = MutableLiveData<List<List<String>>>()
    val data: LiveData<List<List<String>>> get() = _data

    fun readExcelFile(context: Context, uri: Uri) {
        viewModelScope.launch(Dispatchers.IO) {
            val contentResolver = context.contentResolver
            val inputStream: InputStream? = contentResolver.openInputStream(uri)
            val data = mutableListOf<List<String>>()
            val formatter = DataFormatter()

            inputStream.use { stream ->
                val bufferedStream = BufferedInputStream(stream)
                val workbook = WorkbookFactory.create(bufferedStream)
                val sheet = workbook.getSheetAt(0)
                val totalRows = sheet.physicalNumberOfRows - 1
                for ((rowIndex, row) in sheet.drop(1).withIndex()) {
                    val rowData = mutableListOf<String>()
                    for (cell in row) {
                        rowData.add(formatter.formatCellValue(cell))
                    }
                    data.add(rowData)
                    _progress.postValue((rowIndex + 1) * 100 / totalRows)
                }
            }

            _data.postValue(data)
        }
    }

    fun resetProgress() {
        _progress.value = 0
    }

    fun convertToKML(data: List<List<String>>, context: Context): Uri? {
        val kmlContent = buildString {
            append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
            append("<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n")
            append("<Document>\n")
            data.drop(1).forEach { row ->
                append(createPlaceMark(row))
            }
            append("</Document>\n")
            append("</kml>\n")
        }

        return try {
            val file = File(context.cacheDir, "output.kml")
            FileOutputStream(file).use { it.write(kmlContent.toByteArray()) }
            FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", file)
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    private fun createPlaceMark(row: List<String>): String {
        val placeMark = """
                    <Placemark>
                        <name>${row[0]}</name>
                        <description>
                        <![CDATA[
                            <p>
                               CAN: ${row[0]} <br/>
                               Meter Code: ${row[1]} <br/>
                               MRU: ${row[2]} <br/>
                               BA: ${row[3]} <br/>
                               Name: ${row[4]} <br/>
                               Address: ${row[5]} <br/>
                               DMA: ${row[6]} <br/>
                               DMZ: ${row[7]} <br/>
                               Latitude: ${row[9]} <br/>
                               Longitude: ${row[10]} <br/>
                               Coordinates Source: ${row[11]} <br>
                            </p>
                        ]]>
                        </description>
                        <Point>
                            <coordinates>${row[10]},${row[9]}</coordinates>
                        </Point>
                    </Placemark>
        """.trimIndent()

        return placeMark
    }
}