package com.zorenkonte.exceltokml

import android.Manifest
import android.content.Context
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
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Home
import androidx.compose.material.icons.filled.Info
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.livedata.observeAsState
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.ContextCompat
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import com.zorenkonte.exceltokml.ui.theme.ExcelToKMLTheme
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.launch

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
                        MainScreen(excelViewModel)
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
        uri?.let {
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

    ExcelFilePickerUI(
        conversionStarted = conversionStarted,
        progress = progress,
        data = data,
        animatedProgress = animatedProgress,
        showDialog = showDialog,
        onFileClick = {
            handleFileClick(
                context = context,
                viewModel = viewModel,
                filePickerLauncher = filePickerLauncher,
                permissionLauncher = permissionLauncher,
                excelMimeTypes = excelMimeTypes,
                conversionStarted = conversionStarted // Pass the conversionStarted state
            )
        },
        onKMLClick = {
            handleKMLClick(context, viewModel, data)
        }
    )
}

@Composable
fun ExcelFilePickerUI(
    conversionStarted: MutableState<Boolean>,
    progress: Int,
    data: List<List<String>>,
    animatedProgress: Float,
    showDialog: MutableState<Boolean>,
    onFileClick: () -> Unit,
    onKMLClick: () -> Unit
) {
    Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
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

            Row(verticalAlignment = Alignment.CenterVertically) {
                if ((!conversionStarted.value && progress == 0) || progress == 100) {
                    Button(onClick = { onFileClick() }) {
                        Text(text = if (data.isNotEmpty()) "Convert Again" else "Convert")
                    }
                }

                if (data.isNotEmpty() && !(conversionStarted.value && progress == 0) && progress !in 1..99) {
                    Spacer(modifier = Modifier.width(16.dp))
                    Button(onClick = { onKMLClick() }) {
                        Text(text = "Open KML")
                    }
                }
            }

            LazyColumn {
                items(data) { row ->
                    DisplayRow(row)
                }
            }

            if (showDialog.value) {
                PermissionDialog { showDialog.value = false }
            }
        }
    }
}

@Composable
fun DisplayRow(row: List<String>) {
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

@Composable
fun PermissionDialog(onDismiss: () -> Unit) {
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(text = "Permission Denied") },
        text = { Text(text = "Permission to read external storage is required to select an Excel file.") },
        confirmButton = {
            Button(onClick = onDismiss) {
                Text("OK")
            }
        }
    )
}

fun handleFileClick(
    context: Context,
    viewModel: ExcelViewModel,
    filePickerLauncher: androidx.activity.compose.ManagedActivityResultLauncher<Array<String>, Uri?>,
    permissionLauncher: androidx.activity.compose.ManagedActivityResultLauncher<String, Boolean>,
    excelMimeTypes: Array<String>,
    conversionStarted: MutableState<Boolean>
) {
    viewModel.resetProgress()
    conversionStarted.value = false // Reset the conversionStarted state
    val permission = getRequiredPermission()
    if (ContextCompat.checkSelfPermission(
            context,
            permission
        ) == PackageManager.PERMISSION_GRANTED
    ) {
        filePickerLauncher.launch(excelMimeTypes)
    } else {
        permissionLauncher.launch(permission)
    }
}

fun handleKMLClick(context: Context, viewModel: ExcelViewModel, data: List<List<String>>) {
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
            Toast.makeText(context, "No application found to open KML file", Toast.LENGTH_SHORT)
                .show()
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

@Composable
fun MainScreen(viewModel: ExcelViewModel) {
    val drawerState = rememberDrawerState(DrawerValue.Closed)
    val scope = rememberCoroutineScope()
    val navController = rememberNavController()

    ModalNavigationDrawer(
        drawerState = drawerState,
        drawerContent = {
            DrawerContent(navController, drawerState, scope)
        }
    ) {
        NavHost(navController, startDestination = "home") {
            composable("home") { ExcelFilePicker(viewModel) { } }
            composable("credits") { CreditsScreen() }
        }
    }
}

@Composable
fun DrawerContent(
    navController: androidx.navigation.NavHostController,
    drawerState: DrawerState,
    scope: CoroutineScope
) {
    var currentRoute by remember { mutableStateOf<String?>(null) }

    LaunchedEffect(navController) {
        navController.currentBackStackEntryFlow.collect { backStackEntry ->
            currentRoute = backStackEntry.destination.route
        }
    }

    ModalDrawerSheet {
        Column(modifier = Modifier.padding(start = 16.dp, end = 16.dp)) {
            NavigationDrawerItem(
                icon = { Icon(Icons.Default.Home, contentDescription = null) },
                label = { Text("Home", fontSize = 18.sp, fontWeight = FontWeight.Medium) },
                selected = currentRoute == "home",
                onClick = {
                    scope.launch { drawerState.close() }
                    navController.navigate("home")
                }
            )

            Spacer(modifier = Modifier.height(8.dp))

            NavigationDrawerItem(
                icon = { Icon(Icons.Default.Info, contentDescription = null) },
                label = { Text("Credits", fontSize = 18.sp, fontWeight = FontWeight.Medium) },
                selected = currentRoute == "credits",
                onClick = {
                    scope.launch { drawerState.close() }
                    navController.navigate("credits")
                }
            )
        }
    }
}

@Composable
fun CreditsScreen() {
    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement = Arrangement.Center
    ) {
        Text(text = "App Credits", fontSize = 24.sp, fontWeight = FontWeight.Bold)
        Spacer(modifier = Modifier.height(16.dp))
        Text(text = stringResource(R.string.developed_by))
        Text(text = stringResource(R.string.version))
    }
}
