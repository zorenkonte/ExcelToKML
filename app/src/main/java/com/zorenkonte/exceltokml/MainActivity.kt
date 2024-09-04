package com.zorenkonte.exceltokml

import android.Manifest
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.activity.viewModels
import androidx.compose.animation.core.animateFloatAsState
import androidx.compose.animation.core.tween
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.livedata.observeAsState
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
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
        targetValue = progress / 100f,
        animationSpec = tween(durationMillis = 500), label = ""
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
            Text(text = "Convert")
        }

        if (conversionStarted.value && progress == 0) {
            CircularProgressIndicator()
        } else if (progress in 1..99) {
            LinearProgressIndicator(
                progress = { animatedProgress },
            )
            Text(text = "${(animatedProgress * 100).toInt()}%")
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
                        Text(text = "Location: ${row[9]}, ${row[10]}")
                    }
                }
            }
        }

        if (showDialog.value) {
            AlertDialog(
                onDismissRequest = { showDialog.value = false },
                title = { Text(text = "Permission Denied") },
                text = { Text(text = "Permission to read external storage is required to select an Excel file.") },
                confirmButton = {
                    Button(onClick = { showDialog.value = false }) {
                        Text("OK")
                    }
                }
            )
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