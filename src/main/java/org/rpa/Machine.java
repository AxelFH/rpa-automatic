package org.rpa;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.image.BufferedImage;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.Toolkit;
import java.io.IOException;
import java.util.ArrayList;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.File;
import java.util.HashSet;
import java.util.Random;
import java.util.Set;
import org.rpa.Folio;
import org.rpa.ModalFolios;

// --- Proceso ---
// 1. Obtiene nombre y folio de Excel
//  - Si ya fue cargado, continua el proceso
//  - Si no fue cargado, continúa buscando
// 2. Entra a TekAuto y busca el ID
// 3. Descarga el archivo
// 4. Entra a Validoc y crea un nuevo expediente
// 5. Selecciona el archivo, lo carga, y sube el expediente
//  - Detectar el nombre del archivo descargado
//  - Graceful logout Validoc
// 6. Repetir

// --- Consideraciones ---
// 1. Control de proceso (interrupción, pausa, etc.)
// 2. Intentar saltar directamente a row necesaria para continuar proceso
//      - (únicamente al arranque del script)
// 3. Validación contra descarga no finalizada de archivos
// 4. Confirmar realización de acciones con validación segundo a
//    segundo a partir de color extraído de un pixel decisivo en pantalla

public class Machine {
    static Robot bot;

    Robot robot = new Robot();

    ArrayList<Client> clients = new ArrayList<>();
    int traversedRows = 0;

    public Machine() throws AWTException {
        bot = new Robot();
    }

    private static WebDriver driver;

    private static void initializeWebDriver() {
    }

    public void startBot(){
        try {
            openSpreadsheet();
            rowCatchUp("START");
            cycleRows();
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    public void openSpreadsheet() throws Exception {
        // Presiona tecla windows
        bot.keyPress(KeyEvent.VK_WINDOWS);
        bot.keyRelease(KeyEvent.VK_WINDOWS);

        // Espera un segundo
        Thread.sleep(1000);

        // Escribe "CHROME"
        textEntry("CHROME");

        // Espera un segundo
        Thread.sleep(1000);

        // Presiona enter
        bot.keyPress(KeyEvent.VK_ENTER);
        bot.keyRelease(KeyEvent.VK_ENTER);

        // Espera aproximadamente 4 segundos a abrir
        Thread.sleep(5000);
        // Operaciones a realizar si se está abriendo la speadsheet
        // Hacemos clic en el centro de la pantalla para devolver focus
        robot.mouseMove(500, 65);

        robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

        Thread.sleep(250);

        // Escribe URL de la hoja de cálculo
        textEntry("https://santandernet-my.sharepoint.com/:x:/r/personal/z285685_santander_com_mx/Documents/BASE%20ASIGNACION%20DPS/BASE%20ASIGNACION%20NUEVA.xlsx?d=wdcc242eed7824f1690adfcc776c4903d&csf=1&web=1&e=dazdPj");

        // Presiona ENTER
        bot.keyPress(KeyEvent.VK_ENTER);
        bot.keyRelease(KeyEvent.VK_ENTER);

        // Espera 10 segundos a que cargue la hoja
        Thread.sleep(10000);

        // Da clic en coordenadas de primer folio
//        clickAtCoordinates(165, 365);
    }

    public void rowCatchUp(String operation) throws Exception {
        if (operation.equals("START")) {
            // Operaciones a realizar si se está abriendo la speadsheet
            // Hacemos clic en el centro de la pantalla para devolver focus
            robot.mouseMove(683, 384);

            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            Thread.sleep(250);

            bot.keyPress(KeyEvent.VK_CONTROL);
            bot.keyPress(KeyEvent.VK_HOME);
            bot.keyRelease(KeyEvent.VK_CONTROL);
            bot.keyRelease(KeyEvent.VK_HOME);

            Thread.sleep(250);

            // Da clic en coordenadas de primer folio
            clickAtCoordinates(165, 365);

            Thread.sleep(250);

            pressKey(5, KeyEvent.VK_RIGHT);

            String hasValue = getCellValue();

            Thread.sleep(50);

            while (!hasValue.equals("")) {
                pressKey(1, KeyEvent.VK_DOWN);
                traversedRows = traversedRows + 1;

                Thread.sleep(50);

                hasValue = getCellValue();
            }

            pressKey(5, KeyEvent.VK_LEFT);
        }

        if (operation.equals("CONTINUE")) {
            pressKey(1, KeyEvent.VK_DOWN);
            traversedRows = traversedRows + 1;

            Thread.sleep(50);
        }

        pressKey(5, KeyEvent.VK_RIGHT);

        Thread.sleep(50);

        String currentStatus = getCellValue();

        Thread.sleep(50);

        pressKey(5, KeyEvent.VK_LEFT);

        if (currentStatus.equals("")) {
            return;
        } else {
            Thread.sleep(50);

            rowCatchUp("CONTINUE");
        }
    }

    public String checkRow() throws Exception{
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Clipboard clipboard = toolkit.getSystemClipboard();

        // Limpia el portapapeles para comenzar ciclo
        StringSelection emptySelection = new StringSelection("");
        clipboard.setContents(emptySelection, null);

        // Revisa si existe folio, si no existe, toma break de 10 min
        String folio = getCellValue();

        if (folio.equals("")){
            return "SIN_FOLIO";
        }

        // Se desplaza tres columnas a la derecha para posicionarse en agencia
        pressKey(3, KeyEvent.VK_RIGHT);

        // Obtiene el valor de la agencia en mayúsculas para comparar
        String dealership = getCellValue().toUpperCase();
        pressKey(1, KeyEvent.VK_RIGHT);

        // Se desplaza una columna a la derecha para posicionarse en tipo
        String type = getCellValue().toUpperCase();
        pressKey(1, KeyEvent.VK_RIGHT);

        // Se desplaza una columnaa a la derecha para posicionarse en estatus
        String status = getCellValue();

        pressKey(5, KeyEvent.VK_LEFT);

        if (dealership.equals("STAFF SANTA FE") || dealership.equals("SUCURSAL")) {
            return "ABORTADO_AGENCIA";
        } else if (!status.equals("")) {
            return "YA_CARGADO";
        } else if (!type.equals("NUEVO")) {
            return "NO_NUEVO";
        } else {
            return "CONTINUA_CARGA";
        }
    }

    boolean shouldFocus = false;

    public void cycleRows() throws Exception {
        // Revisa el estado de la fila para determinar la acción a seguir
        // "SIN_FOLIO" - No existe folio -> Terminó lectura
        // "ABORTADO_AGENCIA" No aplica por agencia -> Escribe y salta a siguiente fila
        // "YA_CARGADO" - Estatus ya tenía contenido -> Salta a siguiente columna
        // "CONTINUA_CARGA" - No hay errores -> Comienza proceso de carga

        String state = checkRow();

        if (state.equals("SIN_FOLIO")) {
            // Empieza a correr procesos sobre ArrayList de clientes
            // Si no hay elementos restantes en ArrayList, toma descanso
            // Cierra excel y toma break de 5 minutos

//            if (driver != null) {
//                driver.quit();
//            }
            bot.mouseMove(1337, 23);

            Thread.sleep(50);

            bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Espera cinco minutos
            Thread.sleep(300000);
            openSpreadsheet();
            rowCatchUp("START");

            cycleRows();
        } else if (state.equals("ABORTADO_AGENCIA")) {
            // Si agencia dice "STAFF SANTA FE" o "SUCURSAL", el estatus
            // se escribe directamente como "NO APLICA", y salta la carga
            // y avanza a la siguiente fila
            pressKey(5, KeyEvent.VK_RIGHT);
            Thread.sleep(50);

            textEntry("NO APLICA");

            Thread.sleep(50);

            pressKey(5, KeyEvent.VK_LEFT);
            pressKey(1, KeyEvent.VK_DOWN);
            traversedRows = traversedRows + 1;

            Thread.sleep(50);

            cycleRows();
        } else if (state.equals("NO_NUEVO")) {
            pressKey(1, KeyEvent.VK_DOWN);
            traversedRows = traversedRows + 1;
        } else if (state.equals("YA_CARGADO")) {
            pressKey(1, KeyEvent.VK_DOWN);
            traversedRows = traversedRows + 1;
        } else if (state.equals("CONTINUA_CARGA")) {
            // Obtiene los datos necesarios para la carga
            // y avanza a la siguiente columna para seguir recopilando
            String folio = getCellValue();

            pressKey(1, KeyEvent.VK_RIGHT);
            String nombre = getCellValue();

            Client cliente = new Client(folio, nombre);

            clients.add(cliente);
            // Ejecuta process handler y debe regresar
            // a ventana de Excel con foco en la ultima celda trabajada
            // (cliente)

            // Inicia en ventana TekAuto y termina en Excel para continuar
            String processResponse = processHandler(cliente, false);

            Thread.sleep(3000);

            // Se mueve desde nombre de cliente a estatus
            pressKey(5, KeyEvent.VK_RIGHT);
            Thread.sleep(50);

            // Presiona F2 para editar la celda
            bot.keyPress(KeyEvent.VK_F2);
            bot.keyRelease(KeyEvent.VK_F2);
            Thread.sleep(50);

            if (processResponse == "CARGA_EXITOSA") {
                textEntry("DPS");
            } else if (processResponse == "ERROR_ARCHIVO") {
                textEntry("ARCHIVO DAÑADO");
            }

            Thread.sleep(50);

            bot.keyPress(KeyEvent.VK_ENTER);
            Thread.sleep(50);
            bot.keyRelease(KeyEvent.VK_ENTER);

            Thread.sleep(50);
            pressKey(1, KeyEvent.VK_UP);

            Thread.sleep(50);

            pressKey(5, KeyEvent.VK_LEFT);
            Thread.sleep(50);
            pressKey(1, KeyEvent.VK_DOWN);
            traversedRows = traversedRows + 1;
        }

        cycleRows();
    }

    private Set<String> listFilesInFolder() {
        Set<String> files = new HashSet<>();
        File folder = new File("C:\\Users\\Z285641\\Downloads\\");
        File[] listOfFiles = folder.listFiles();

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile()) {
                    files.add(file.getName());
                }
            }
        }

        return files;
    }

    public String checkForNewFile(Set<String> previousFileList) {
        String folderPath = "C:\\Users\\Z285641\\Downloads\\";

        Set<String> currentFileList = listFilesInFolder();

        for (String file : currentFileList) {
            if (!previousFileList.contains(file)) {

                return folderPath + file;
            }
        }

        return null; // No new file found
    }

    public String processHandler(Client cliente, boolean shouldFocus) throws AWTException, InterruptedException {
        // Configura ubicación de carpeta de descargas
        String folderPath = "C:\\Users\\Z285641\\Downloads\\";

        // Configura la ruta del EdgeDriver
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Z285641\\Downloads\\2\\chromedriver-win64\\chromedriver.exe");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("start-maximized");
        driver = new ChromeDriver(options);

        try {
            WebDriverWait wait = new WebDriverWait(driver, 20);

            // Abre la página web
            driver.get("https://creditoautobo.mx.corp/tekfinauto/tekorigination/default.aspx");

            // Espera a que la página cargue
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(5000);

            // Encuentra el campo de usuario y escribe en él
            WebElement userField = driver.findElement(By.id("txtLogin"));
            userField.sendKeys("X318428");

            // Encuentra el campo de contraseña y escribe en él
            WebElement passwordField = driver.findElement(By.id("txtPassword"));
            passwordField.sendKeys("Documentos04");

            // Encuentra el botón de login y haz clic en él
            WebElement loginButton = driver.findElement(By.id("btnLogin"));
            loginButton.click();

            // Espera a que la página de inicio cargue después del inicio de sesión
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(5000);

            // Navega a la nueva URL dentro del mismo sitio
            driver.get("https://creditoautobo.mx.corp/tekfinauto/tekorigination//CustomerASPx/wbCustomerCenterResumeBPM.aspx?PersonId=@1&cntPrmA=X318428&cntPrmB=5&iOpReference=0&iIdReference=0&customerCode=&searchCriteria=&path=&impoundIdOperation=0");

            // Espera a que la nueva página cargue completamente
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(5000);

            // 24. Escribe folio en campo ID
            WebElement searchIdField = driver.findElement(By.id("txiSPersonId"));
            searchIdField.sendKeys(cliente.folio);

            // Dar clic en botón buscar
            WebElement btnSearchIdField = driver.findElement(By.id("ctl00_BodyContent_btnSearch"));
            btnSearchIdField.click();

            Thread.sleep(3000);

            robot.mouseMove(683, 384);

            // Optionally, click to focus on the specific area
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Scroll down in the area (positive value for down, negative for up)
            robot.mouseWheel(3);

            // Scrolls down 5 notches

            Thread.sleep(3000);

            robot.mouseMove(90, 523);

            // Optionally, click to focus on the specific area
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // 28. Accionar clic "Acciones sobre la solicitud de crédito" en botón "Expediente"
            //   - ctl00_BodyContent_btcCreditformCustomAction1
            WebElement btnExpedient = driver.findElement(By.id("ctl00_BodyContent_btcCreditformCustomAction1"));
            btnExpedient.click();

            Thread.sleep(15000);

            // 31. Selecciona "Expediente para evaluación titular"
            //   - Busca elemento cuyo contenido sea "EXPEDIENTE PARA EVALUACION - TITULAR"
//                 clickAtCoordinates(150, 362);
            bot.mouseMove(150, 362);
            bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
//
            Thread.sleep(7500);

            // 32. Espera a que se visualice botón y acciona "Visor de documentos"
//                WebElement btnDocumentViewer = driver.findElement(By.id("btnDocumentViewer"));
//                btnDocumentViewer.click();
            bot.mouseMove(140, 568);
            bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
//
            Thread.sleep(7500);

            // PENDIENTE: Obtener archivos en carpeta de descargas
            Set<String> previousFileList = listFilesInFolder();

            // Descarga documento
            // 34. Accionar botón "Abrir" en visor de documentos
//                WebElement btnDocumentDownload = driver.findElement(By.id("open-button"));
//                btnDocumentDownload.click();
            bot.mouseMove(370, 450);

            // PENDIENTE: Buscar buscar cambios en carpeta de descargas
            String filePath = checkForNewFile(previousFileList);

            // Después de dos segundos, revisar en intervalos de un segundo
            // si se descargó un archivo, en caso de existir, revisa que la
            // última extensión no sea "crdownload" para verificar que se esté
            // descargando

            int downloadIntents = 3;
            int timeToStart = 10;
            int timeToEnd = 10;
            boolean downloaded = false;
            boolean started = false;

            // Revisa si el tamaño del archivo está cambiando, cada que
            // el tamaño cambia (continúa aumentando), el contador se
            // reinicia a cero; cada tick (segundo) que no se detecte un
            // cambio en el tamaño del archivo, o que la última extensión
            // ya no sea "crdownload"; considera que el archivo ya terminó
            // de ser descargado y procede
            do {
                // Espera 1 segundo antes de iniciar el ciclo de validación
                // de descarga
                Thread.sleep(1000);
                downloadIntents = downloadIntents - 1;

                bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

                do {
                    // Revisa si filePath es nulo para confirmar que no hay nuevos
                    // archivos en la carpeta de descargas
                    filePath = checkForNewFile(previousFileList);
                    started = filePath != null;

                    timeToStart = timeToStart - 1;
                    Thread.sleep(1000);
                } while (timeToStart > 0 && !started);

                do {
                    downloaded = isPdfFile(filePath);

                    if (!downloaded) {

                        if (timeToEnd < 5) {
                            Random random = new Random();
                            if (random.nextInt(3) == 0) {

                                bot.keyPress(KeyEvent.VK_CONTROL);
                                bot.keyPress(KeyEvent.VK_J);
                                bot.keyRelease(KeyEvent.VK_CONTROL);
                                bot.keyRelease(KeyEvent.VK_J);

                                Thread.sleep(3000);

                                clickAtCoordinates(975, 275);
                                Thread.sleep(1000);
                                clickAtCoordinates(970, 270);

                                Thread.sleep(3000);
                                bot.keyPress(KeyEvent.VK_CONTROL);
                                bot.keyPress(KeyEvent.VK_W);
                                bot.keyRelease(KeyEvent.VK_CONTROL);
                                bot.keyRelease(KeyEvent.VK_W);


                                timeToEnd = 10;
                            }
                        } else {
                            timeToEnd = timeToEnd - 1;
                        }

                    }
                    Thread.sleep(1000);
                } while (!downloaded);
            } while (downloadIntents > 0 && !downloaded);

            if (!downloaded) {
                return "ERROR_ARCHIVO";
            }

            // 35. Irse a validoc (https://validoc.mx)
            driver.get("https://validoc.mx");
//
//                // Espera a que la nueva página cargue completamente
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(5000);
//
//                //   - Ingresa en validoc
            WebElement validocUsrField = driver.findElement(By.id("UserName"));
            validocUsrField.sendKeys("GrupoApa_uautomatico");
//
            WebElement validocPwdField = driver.findElement(By.id("Password"));
            validocPwdField.sendKeys("4tOM4T&)");
//
//                // Da clic en botón de inicio de sesión
            WebElement submitButton = driver.findElement(By.xpath("//input[@type='submit']"));
            submitButton.click();
//
//                // Espera a que la página de inicio cargue después del inicio de sesión
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(5000);

            //   - Revisar tema de login/logout

            // 37. Pestaña expedientes en sidebar https://validoc.mx/Expedient
            driver.get("https://validoc.mx/Expedient");
//
//                // Espera la página
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("html")));
            Thread.sleep(10000);

            // 38. Botón "nuevo", lateral superior derecho btnChangeStage
            WebElement validocCreateExpedient = driver.findElement(By.id("btnCreateExpedient"));
            validocCreateExpedient.click();
//
            Thread.sleep(3000);
//
//                // 39. Dropdown: "Auto" TypeOperationId
            WebElement validocTypeOperation = driver.findElement(By.id("TypeOperationId"));
            Select operationTypeDropdown = new Select(validocTypeOperation);
            operationTypeDropdown.selectByIndex(1);
//
//                // 43. Dropdown: "Completo"
            WebElement validocTypeExpedient = driver.findElement(By.id("TypeExpedient"));
            Select expedientTypeDropdown = new Select(validocTypeExpedient);
            expedientTypeDropdown.selectByIndex(1);
//
//                // 45. Escribir folio (estático)
            WebElement validocFolioField = driver.findElement(By.id("Folio"));
            validocFolioField.sendKeys(cliente.folio);
//
//                // 47. Escribir cliente (estático))
            WebElement validocClientField = driver.findElement(By.id("Accredited"));
            validocClientField.sendKeys(cliente.nombre);

            Thread.sleep(3000);

            // Accionar botón de crear
            bot.mouseMove(885, 595);
            bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            Thread.sleep(10000);

            // VALIDOC EXP --- 1. Seleccionar archivo opFileInput_2
            WebElement fileInput = driver.findElement(By.id("opFileInput_2"));
            fileInput.sendKeys(filePath);

            // VALIDOC EXP --- 2. Clic en nubecita <a class="btn" onclick="botonUpload(&quot;opFileInput_2&quot;, &quot;0c819272-85fc-4ccf-b245-646fed1ac0be&quot;, &quot;2&quot;)"><i class="fa fa-cloud-upload fa-2x" style="color:blue"></i></a>
            // Use AWT Robot to move the mouse to the specific coordinates (e.g., 683, 384)
            robot.mouseMove(683, 384);

            // Optionally, click to focus on the specific area
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Scroll down in the area (positive value for down, negative for up)
            robot.mouseWheel(4); // Scrolls down 4 notches
            robot.mouseMove(860, 395);

            Thread.sleep(3000);

            // Optionally, click to focus on the specific area
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Revisar el color de la pantalla para saber si ya terminó de cargar
            // el archivo

            String hex;

            Thread.sleep(2000);

            // #ffffff - Terminado
            // #7f7f7f - Carga en progreso
            do {
                hex = getPixelHexColor(1100, 400);
                System.out.printf("Hex: %s\n", hex);

                Thread.sleep(1000);
            } while (hex.equals("#7f7f7f"));

            Thread.sleep(3000);

            hex = getPixelHexColor(1200, 705);

            Thread.sleep(1000);

            if (hex.equals("#f1eff1")) {
                robot.mouseMove(1320, 705);
                Thread.sleep(100);
                robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

                Thread.sleep(1000);

                robot.mouseMove(1100, 400);
                Thread.sleep(100);
                robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }

            // Espera a que se suba el archivo
            // Thread.sleep(45000);

            // Regresa al inicio de la página
            robot.mouseWheel(-10); // Scrolls up 10 notches
            Thread.sleep(2000);

            // Scroll down in the area (positive value for down, negative for up)
            robot.mouseWheel(7); // Scrolls down 5 notches
            robot.mouseMove(775, 570);

            Thread.sleep(2000);

            // Optionally, click to focus on the specific area
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(100);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Revisar el color de la pantalla para saber si ya terminó de cargar
            // el archivo

            // #ffffff - Terminado
            // #7f7f7f - Carga en progreso

            Thread.sleep(2000);

            do {
                hex = getPixelHexColor(1100, 400);
                System.out.printf("Hex: %s\n", hex);

                Thread.sleep(1000);
            } while (hex.equals("#7f7f7f"));

            Thread.sleep(3000);

            hex = getPixelHexColor(1200, 705);

            Thread.sleep(1000);

            if (hex.equals("#f1eff1")) {
                robot.mouseMove(1320, 705);
                Thread.sleep(100);
                robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

                Thread.sleep(1000);

                robot.mouseMove(1100, 400);
                Thread.sleep(100);
                robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }

            // Espera a que se suba el archivo
            // Thread.sleep(45000);

            robot.mouseMove(1000, 570);
            Thread.sleep(100);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            robot.mouseWheel(-10); // Scrolls up 10 notches

            // Posiciona el cursor en el botón de cierre de sesión
            robot.mouseMove(1328, 166);

            // Espera dos segundos en caso de bloqueos de cursor
            Thread.sleep(2000);

            // Da clic para cerrar sesión
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Manual - Asegurar de cerrar sesion y
            // posiblemente cerrar navegador webdriver
            Thread.sleep(5000);

            // Dar clic en cerrar navegador webdriver
            // Posiciona el cursor en el botón de cierre de webDriver
            // robot.mouseMove(1337, 23);
            driver.quit();

            Thread.sleep(5000);

            robot.mouseMove(683, 384);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

            // Presiona tecla WINDOWS + 1 para regresar a Excel
            // donde debe tener foco en la última celda trabajada
//            bot.keyPress(KeyEvent.VK_WINDOWS);
//            bot.keyPress(KeyEvent.VK_1);
//            bot.keyRelease(KeyEvent.VK_WINDOWS);
//            bot.keyRelease(KeyEvent.VK_1);

            Thread.sleep(1000);

            bot.keyPress(KeyEvent.VK_CONTROL);
            bot.keyPress(KeyEvent.VK_HOME);
            Thread.sleep(100);
            bot.keyRelease(KeyEvent.VK_CONTROL);
            bot.keyRelease(KeyEvent.VK_HOME);

            pressKey(1, KeyEvent.VK_RIGHT);
            pressKey(1, KeyEvent.VK_DOWN);

            Thread.sleep(1000);
            pressKey(traversedRows, KeyEvent.VK_DOWN);

            return "CARGA_EXITOSA";
        } catch (NoSuchElementException | InterruptedException e) {
            throw new RuntimeException(e);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    class Client {
        public String folio;
        public String nombre;

        Client(String folio, String nombre){
            this.folio = folio;
            this.nombre = nombre;
        }
    }

    public static boolean isPdfFile(String filePath) {
        if (filePath == null || filePath.isEmpty()) {
            return false;
        }

        String lowerCaseFilePath = filePath.toLowerCase();
        return lowerCaseFilePath.endsWith(".pdf");
    }

    public String getCellValue() throws AWTException, UnsupportedFlavorException, IOException, InterruptedException {
        // Presiona F2 para editar la celda
        bot.keyPress(KeyEvent.VK_F2);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_F2);
        Thread.sleep(25);

        // Presiona Ctrl + A para seleccionar
        bot.keyPress(KeyEvent.VK_CONTROL);
        Thread.sleep(25);
        bot.keyPress(KeyEvent.VK_A);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_A);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_CONTROL);
        Thread.sleep(25);

        // Presiona Ctrl + C para copiar el texto al portapapeles
        bot.keyPress(KeyEvent.VK_CONTROL);
        Thread.sleep(25);
        bot.keyPress(KeyEvent.VK_C);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_C);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_CONTROL);
        Thread.sleep(25);

        // Obtiene el contenido del portapapeles
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Clipboard clipboard = toolkit.getSystemClipboard();
        Thread.sleep(50);

        // Guarda el contenido del portapapeles en "cellValue"
        String cellValue = (String) ((Clipboard) clipboard).getData(DataFlavor.stringFlavor);
        Thread.sleep(100);

        // Una vez guardado el contenido, establece el contenido del portapapeles en una cadena vacía
        StringSelection emptySelection = new StringSelection("");
        clipboard.setContents(emptySelection, null);
        Thread.sleep(100);

        bot.keyPress(KeyEvent.VK_ESCAPE);
        Thread.sleep(25);
        bot.keyRelease(KeyEvent.VK_ESCAPE);

        return cellValue;
    }

    public void robotProcess() throws InterruptedException {
        bot.keyPress(KeyEvent.VK_WINDOWS);
        bot.keyRelease(KeyEvent.VK_WINDOWS);
        Thread.sleep(1000);

        textEntry("CHROME");
        Thread.sleep(1000);

        bot.keyPress(KeyEvent.VK_ENTER);
        bot.keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(4000);

        textEntry("https://santandernet-my.sharepoint.com/:x:/r/personal/z285685_santander_com_mx/_layouts/15/Doc.aspx?sourcedoc=%7BDCC242EE-D782-4F16-90AD-FCC776C4903D%7D&file=BASE%20ASIGNACION%20NUEVA.xlsx&wdLOR=cC03EEE03-A2C8-40D0-A3A9-F650C73A97B6&fromShare=true&action=default&mobileredirect=true");

        Thread.sleep(1000);
        bot.keyPress(KeyEvent.VK_ENTER);
        bot.keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(1000);
    }

    public void pressKey(int times, int key) throws InterruptedException{
        for (int k=0; k<times; k++){
            bot.keyPress(key);
            bot.keyRelease(key);
            Thread.sleep(50);
        }
    }
    public void timesTab(int times) throws InterruptedException {
        for (int j = 0; j < times; j++) {
            bot.keyPress(KeyEvent.VK_TAB);
            bot.keyRelease(KeyEvent.VK_TAB);
            Thread.sleep(500);
        }
    }

    public void clickAtCoordinates(int x, int y) throws Exception {
        bot.mouseMove(x, y);
        bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void textEntry(String text) {
        for (char c : text.toCharArray()) {
            if (Character.isUpperCase(c)) {
                bot.keyPress(KeyEvent.VK_SHIFT);
                int keyCode = KeyEvent.getExtendedKeyCodeForChar(Character.toLowerCase(c));
                bot.keyPress(keyCode);
                bot.keyRelease(keyCode);
                bot.keyRelease(KeyEvent.VK_SHIFT);
            } else {
                if (c == '&' || c == '%' || c == '!' || c == '=' || c == ':'
                        || c == '/' || c == ';' || c == '_' || c == '?' || c == '\"') {
                    switch (c) {
                        case '&':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_6);
                            bot.keyRelease(KeyEvent.VK_6);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '%':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_5);
                            bot.keyRelease(KeyEvent.VK_5);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '!':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_1);
                            bot.keyRelease(KeyEvent.VK_1);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '=':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_0);
                            bot.keyRelease(KeyEvent.VK_0);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case ':':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_PERIOD);
                            bot.keyRelease(KeyEvent.VK_PERIOD);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '/':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_7);
                            bot.keyRelease(KeyEvent.VK_7);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '\"':
                            bot.keyPress(KeyEvent.VK_BACK_SLASH);
                            bot.keyRelease(KeyEvent.VK_BACK_SLASH);
                            break;
                        case ';':
                            bot.keyPress(KeyEvent.VK_SEMICOLON);
                            bot.keyRelease(KeyEvent.VK_SEMICOLON);
                            break;
                        case '?':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_QUOTE);
                            bot.keyRelease(KeyEvent.VK_QUOTE);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '_':
                            bot.keyPress(KeyEvent.VK_SHIFT);
                            bot.keyPress(KeyEvent.VK_MINUS);
                            bot.keyRelease(KeyEvent.VK_SHIFT);
                            bot.keyRelease(KeyEvent.VK_MINUS);
                            break;
                        default:
                            System.out.println("Character " + c + " not handled.");
                            break;
                    }
                } else {
                    int keyCode = KeyEvent.getExtendedKeyCodeForChar(c);
                    bot.keyPress(keyCode);
                    bot.keyRelease(keyCode);
                }
            }
            bot.delay(50);
        }
    }

    private static String getPixelHexColor(int x, int y) throws AWTException {
        // Create a Robot instance
        Robot robot = new Robot();

        // Capture the screen size
        Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());

        // Take a screenshot and get the BufferedImage
        BufferedImage screenFullImage = robot.createScreenCapture(screenRect);

        // Get the color of the pixel at (x, y)
        Color pixelColor = new Color(screenFullImage.getRGB(x, y));

        // Dispose of the BufferedImage to free up memory
        screenFullImage.flush();

        // Convert the color to hex format
        return String.format("#%02x%02x%02x", pixelColor.getRed(), pixelColor.getGreen(), pixelColor.getBlue());
    }

    public Robot getBot() {
        return bot;
    }

    public void setBot(Robot bot) {
        this.bot = bot;
    }
}