package com.zorenkonte.exceltokml

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.activity.viewModels
import androidx.compose.animation.core.animateFloatAsState
import androidx.compose.animation.core.tween
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.LinearProgressIndicator
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.livedata.observeAsState
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.ContextCompat
import com.zorenkonte.exceltokml.ui.theme.ExcelToKMLTheme

class MainActivity : ComponentActivity() {
    private val excelViewModel: ExcelViewModel by viewModels()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            ExcelToKMLTheme {
                Scaffold(modifier = Modifier.fillMaxSize()) { innerPadding ->
                    Box(
                        contentAlignment = Alignment.Center,
                        modifier = Modifier
                            .fillMaxSize()
                            .padding(innerPadding)
                    ) {
                        ExcelFilePicker(excelViewModel) { uri ->
                            println(uri)
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun ExcelFilePicker(viewModel: ExcelViewModel, onFileSelected: (Uri?) -> Unit) {
    val context = LocalContext.current
    val showDialog = remember { mutableStateOf(false) }
    val progress by viewModel.progress.observeAsState(0)
    val data by viewModel.data.observeAsState(emptyList())
    val conversionStarted = remember { mutableStateOf(false) }

    val animatedProgress by animateFloatAsState(
        targetValue = progress / 100f, animationSpec = tween(durationMillis = 500), label = ""
    )

    val excelMimeTypes = arrayOf(
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    val filePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        if (uri != null) {
            viewModel.readExcelFile(context, uri)
            conversionStarted.value = true
        }
        onFileSelected(uri)
    }

    val permissionLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.RequestPermission()
    ) { isGranted: Boolean ->
        if (isGranted) {
            filePickerLauncher.launch(excelMimeTypes)
        } else {
            showDialog.value = true
        }
    }

    Column(horizontalAlignment = Alignment.CenterHorizontally) {
        if (conversionStarted.value && progress == 0) {
            CircularProgressIndicator()
            Spacer(modifier = Modifier.height(8.dp))
        } else if (progress in 1..99) {
            LinearProgressIndicator(
                progress = { animatedProgress },
            )
            Spacer(modifier = Modifier.height(8.dp))
            Text(
                text = "${(animatedProgress * 100).toInt()}%",
                fontSize = 28.sp,
                fontWeight = FontWeight(500)
            )
        }

        Row(
            verticalAlignment = Alignment.CenterVertically
        ) {
            if ((!conversionStarted.value && progress == 0) || progress == 100) {
                Button(onClick = {
                    viewModel.resetProgress()
                    conversionStarted.value = false
                    val permission = getRequiredPermission()
                    when (PackageManager.PERMISSION_GRANTED) {
                        ContextCompat.checkSelfPermission(context, permission) -> {
                            filePickerLauncher.launch(excelMimeTypes)
                        }

                        else -> {
                            permissionLauncher.launch(permission)
                        }
                    }
                }) {
                    Text(text = if (data.isNotEmpty()) "Convert Again" else "Convert")
                }
            }

            if (data.isNotEmpty() && !(conversionStarted.value && progress == 0) && progress !in 1..99) {
                Spacer(modifier = Modifier.width(16.dp))
                Button(onClick = {
                    val kmlUri = viewModel.convertToKML(data, context)
                    kmlUri?.let {
                        val intent = Intent(Intent.ACTION_VIEW).apply {
                            setDataAndType(it, "application/*")
                            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                        }
                        val chooser = Intent.createChooser(intent, "Open KML file with")
                        if (intent.resolveActivity(context.packageManager) != null) {
                            context.startActivity(chooser)
                        } else {
                            // Handle the case where no application can handle the intent
                            Toast.makeText(
                                context, "No application found to open KML file", Toast.LENGTH_SHORT
                            ).show()
                        }
                    }
                }) {
                    Text(text = "Open KML")
                }
            }
        }

        LazyColumn {
            items(data) { row ->
                Card(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(8.dp),
                    elevation = CardDefaults.cardElevation(2.dp)
                ) {
                    Column(modifier = Modifier.padding(16.dp)) {
                        Text(text = "CAN: ${row[0]}")
                        Text(text = "Name: ${row[4]}")
                        Text(text = "Address: ${row[5]}")
                        Text(text = "Street: ${row[6]}")
                        Text(text = "Location: ${row[10]}, ${row[9]}")
                    }
                }
            }
        }

        if (showDialog.value) {
            AlertDialog(onDismissRequest = { showDialog.value = false },
                title = { Text(text = "Permission Denied") },
                text = { Text(text = "Permission to read external storage is required to select an Excel file.") },
                confirmButton = {
                    Button(onClick = { showDialog.value = false }) {
                        Text("OK")
                    }
                })
        }
    }
}

// Function to get the required permission based on Android version
fun getRequiredPermission(): String {
    return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
        Manifest.permission.READ_MEDIA_IMAGES
    } else {
        Manifest.permission.READ_EXTERNAL_STORAGE
    }
}