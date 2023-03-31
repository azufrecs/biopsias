<?php
include("class/security/index.php");
include("conn/conn.php");
require 'vendor/php-office/autoload.php';
require_once 'Vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpWord\TemplateProcessor;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

/* Limpiando el directorio de informes anteriores */
$MENSAJE = "";
$DIRECTORIO = "subidas/";
$HANDLE = opendir($DIRECTORIO);
while ($FILE = readdir($HANDLE)) {
    if ($FILE != "." && $FILE != ".." && $FILE != ".htaccess" && $FILE != ".gitkeep") {
        unlink($DIRECTORIO . $FILE);
    }
}

$CUENTA_AGREGADOS = 0;
$CUENTA_NO_AGREGADOS = 0;

/* Obteniendo lista de resultados */
$RESULTADOS     = $mysqli->query("SELECT * FROM tbl_biopsias ORDER BY paciente");

/* Procedimiento para importar excel a bd */ 
if (isset($_POST["import"])) {
    $allowedFileType = ['text/xlsx', 'text/xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    if (in_array($_FILES["file"]["type"], $allowedFileType)) {
        $archivos = 'subidas/' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $archivos);
        $spreadsheet = $reader->load($archivos);
        $spreadsheet->setActiveSheetIndex(0);
        $numerofila = $spreadsheet->getActiveSheet()->getHighestRow();

        $COMPROBACION1 = $spreadsheet->getActiveSheet()->getCell('B2')->getValue();
        $COMPROBACION2 = $spreadsheet->getActiveSheet()->getCell('G2')->getValue();
        $COMPROBACION3 = $spreadsheet->getActiveSheet()->getCell('L2')->getValue();
        if ($COMPROBACION1 == "frecep" and $COMPROBACION2 == "organo" and $COMPROBACION3 == "diagnostico") {

            for ($i = 3; $i <= $numerofila; $i++) {
                $data_ano           = $spreadsheet->getActiveSheet()->getCell('A' . $i)->getValue();
                $data_frecep        = $spreadsheet->getActiveSheet()->getCell('B' . $i)->getValue();
                $data_noinforme     = $spreadsheet->getActiveSheet()->getCell('C' . $i)->getValue();
                $data_boriginal     = $spreadsheet->getActiveSheet()->getCell('D' . $i)->getValue();
                $data_laminas       = $spreadsheet->getActiveSheet()->getCell('E' . $i)->getValue();
                $data_bloques       = $spreadsheet->getActiveSheet()->getCell('F' . $i)->getValue();
                $data_organo        = $spreadsheet->getActiveSheet()->getCell('G' . $i)->getValue();
                $data_paciente      = $spreadsheet->getActiveSheet()->getCell('H' . $i)->getValue();
                $data_cid_paciente  = $spreadsheet->getActiveSheet()->getCell('I' . $i)->getValue();
                $data_hospital      = $spreadsheet->getActiveSheet()->getCell('J' . $i)->getValue();
                $data_provincia     = $spreadsheet->getActiveSheet()->getCell('K' . $i)->getValue();
                $data_diagnostico   = $spreadsheet->getActiveSheet()->getCell('L' . $i)->getValue();
                $data_especialista  = $spreadsheet->getActiveSheet()->getCell('M' . $i)->getValue();

                $check = mysqli_query($mysqli, "SELECT * FROM tbl_biopsias WHERE frecep	 = '$data_frecep' AND boriginal = '$data_boriginal' AND organo = '$data_organo' AND paciente = '$data_paciente' AND diagnostico	= '$data_diagnostico' AND especialista = '$data_especialista'");

                if (mysqli_num_rows($check) > 0) {
                    $CUENTA_NO_AGREGADOS++;
                } else {
                    $resultados = mysqli_query($mysqli, "INSERT INTO tbl_biopsias(ano, frecep, noinforme, boriginal, laminas, bloques, organo, paciente, cid_paciente, hospital, provincia, diagnostico, especialista) VALUES('$data_ano', '$data_frecep', '$data_noinforme', '$data_boriginal', '$data_laminas', '$data_bloques', '$data_organo', '$data_paciente', '$data_cid_paciente', '$data_hospital', '$data_provincia', '$data_diagnostico', '$data_especialista')");
                    $CUENTA_AGREGADOS++;
                }

                $MENSAJE = $MENSAJE = "<div class='alert alert-success alert-dismissible fade show' role='alert'><strong>¡Correcto!</strong>&nbsp;Excel procesado satisfactoriamente: " . $CUENTA_AGREGADOS . ' registros agregados de ' . $CUENTA_NO_AGREGADOS . "</div>";
            }
        } else {
            $MENSAJE = "<div class='alert alert-warning alert-dismissible fade show' role='alert'><strong>¡Error!</strong>&nbsp;El fichero que intenta importar no contiene la estructura esperada</div>";
        }
    } else {
        $MENSAJE = "<div class='alert alert-danger alert-dismissible fade show' role='alert'><strong>¡Error!</strong>&nbsp;No ha seleccionado un archivo Excel con extensi&oacute;n XLSX. Por favor vuelva a intentarlo</div>";
    }
}

/* Procedimiento para exportar paciente a Word */ 
if (isset($_POST["export"])) {
    /* Limpiando directorio Exports */
    $DIRECTORIO = "exports/";
    $HANDLE = opendir($DIRECTORIO);
    while ($FILE = readdir($HANDLE)) {
        if ($FILE != "." && $FILE != ".." && $FILE != ".htaccess" && $FILE != ".gitkeep") {
            unlink($DIRECTORIO . $FILE);
        }
    }

    $templateWord = new TemplateProcessor(dirname(__FILE__) . "/template/template.docx");

    /* Obtenemos valores de las variables */
    $PACIENTE   = mysqli_real_escape_string($mysqli, (strip_tags(strtoupper($_POST["cboPaciente"]), ENT_QUOTES)));
    $pieces = explode("|", $PACIENTE);
    $piece_ano          = $pieces['0'];
    $piece_frecep       = $pieces['1'];
    $piece_noinforme    = $pieces['2'];
    $piece_boriginal    = $pieces['3'];
    $piece_laminas      = strlen($pieces['4']) > 0 ? $pieces['4'] : "-";
    $piece_bloques      = strlen($pieces['5']) > 0 ? $pieces['5'] : "-";
    $piece_organo       = $pieces['6'];
    $piece_paciente     = $pieces['7'];
    $piece_cid_paciente = strlen($pieces['8']) > 0 ? $pieces['8'] : "-";
    $piece_hospital     = strlen($pieces['9']) > 0 ? $pieces['9'] : "-";
    $piece_provincia    = strlen($pieces['10']) > 0 ? $pieces['10'] : "-";
    $piece_diagnostico  = strlen($pieces['11']) > 0 ? $pieces['11'] : "-";
    $piece_especialista = strlen($pieces['12']) > 0 ? $pieces['12'] : "-";

    /* Asignamos valores de las variables a la plantilla */
    $templateWord->setValue("biopsia_numero", "CR" . $piece_ano . $piece_noinforme);
    $templateWord->setValue("biopsia_original", $piece_boriginal);
    $templateWord->setValue("organo", $piece_organo);
    $templateWord->setValue("fecha_recepcion", $piece_frecep);
    $templateWord->setValue("fecha_entrega", "--");
    $templateWord->setValue("nombre_paciente", $piece_paciente);
    $templateWord->setValue("cid_paciente", $piece_cid_paciente);
    $templateWord->setValue("hospital", $piece_hospital);
    $templateWord->setValue("provincia", $piece_provincia);
    $templateWord->setValue("material_recibido", $piece_laminas . " LAMINA(S), " .  $piece_bloques . " BLOQUE(S)");
    $templateWord->setValue("diagnostico", $piece_diagnostico);
    $templateWord->setValue("patologo", $piece_especialista);

    $templateWord->saveAs("exports/". $piece_paciente.".docx");

    /* Guardamos el documento */
    if (file_exists("exports/". $piece_paciente.".docx")) {
        header("Content-Description: File Transfer");
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Disposition: attachment; filename=" . basename("exports/". $piece_paciente.".docx"));
        header("Content-Transfer-Encoding: binary");
        header("Expires: 0");
        header("Cache-Control: must-revalidate");
        header("Pragma: public");
        header("Content-Length: " . filesize("exports/". $piece_paciente.".docx"));
        ob_clean();
        flush();
        readfile("exports/". $piece_paciente.".docx");
        exit;
    } else {
        echo "Informe no disponible";
    }
}
?>

<!DOCTYPE html>
<html lang="es">

<head>
    <!-- Etiquetas <meta> obligatorias para Bootstrap -->
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="icon" href="assets/img/favicon.svg">
    <title>Informes Biopsias</title>

    <!-- Enlazando el CSS de Bootstrap -->
    <link href="assets/css/bootstrap.css" rel="stylesheet" media="screen">
    <link href="assets/css/main.css" rel="stylesheet" media="screen">
    <link href="assets/css/fontawesome.css" rel="stylesheet" media="screen">
    <!-- Enlazando el CSS de Bootstrap -->
    <!-- Opcional: enlazando el JavaScript de Bootstrap -->
    <script src="assets/js/jquery-3.6.4.js"></script>
    <script src="assets/js/popper.js"></script>
    <script src="assets/js/bootstrap.js"></script>
    <script src="assets/js/fontawesome.js"></script>
    <!-- Opcional: enlazando el JavaScript de Bootstrap -->
    <script>
        // AUTO HIDE ALERTS
        window.setTimeout(function() {
            $(".alert").fadeTo(500, 0).slideUp(500, function() {
                $(this).remove();
            });
        }, 6000);
    </script>
</head>

<body>
    <!-- Begin page content -->
    <div class="container" align="center">
        <div align="center"><i style="font-size:180px" class="fas fa-microscope text-dark"></i></div>
        <div align="center" class="text-dark" style="font-size:28px">Resultados de Biopsias</div><br>

        <div class="row">
            <div class="col-md-6">
                <div class="card bg-success-30 border-success mb-3">
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center"><i class="fas fa-file-excel fa-6x text-success"></i></div>
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center" class="text-success" style="font-size:22px">Secci&oacute;n para importar fichero Excel con las Biopsias</div>
                    <div class="card-body">
                        <form action='' method='post' name='frmExcelImport' id='frmExcelImport' enctype='multipart/form-data'>
                            <div class='row'>
                                <div class='col-md-8'>
                                    <input type='file' name='file' id='file' enctype='multipart/form-data' class='form-control' accept='.xlsx' id='formFile'>
                                </div>
                                <div class='col-md-4'>
                                    <button type='submit' id='submit' name='import' class='btn btn-success w-100'><i class='fas fa-file-import'></i>&nbsp;&nbsp;Importar Excel</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>

            <div class="col-md-6">
                <div class="card bg-danger-30 border-danger mb-3">
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center"><i class="fas fa-file-word fa-6x text-danger"></i></div>
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center" class="text-danger" style="font-size:22px">Secci&oacute;n un nombre para exportar informe a Word</div>
                    <div class="card-body">
                        <form action='' method='post' name='frmExcelGenerateWord' id='frmExcelGenerateWord' enctype='multipart/form-data'>
                            <div class='row'>
                                <div class='col-md-8'>
                                    <select class="form-select" name="cboPaciente" id="responsive_text" aria-label="Floating label select example" required>
                                        <option disabled value="" selected hidden>Seleccione el Paciente</option>
                                        <?php
                                        while ($rowResultados = $RESULTADOS->fetch_assoc()) {
                                            echo "<option style='white-space:nowrap; text-overflow:elipsis; overflow:hidden;' value='" . $rowResultados['ano'] . "|" . $rowResultados['frecep'] . "|" . $rowResultados['noinforme'] . "|" . $rowResultados['boriginal'] . "|" . $rowResultados['laminas'] . "|" . $rowResultados['bloques'] . "|" . $rowResultados['organo'] . "|" . $rowResultados['paciente'] . "|" . $rowResultados['cid_paciente'] . "|" . $rowResultados['hospital'] . "|" . $rowResultados['provincia'] . "|" . $rowResultados['diagnostico'] . "|" . $rowResultados['especialista'] . "'>" . strtoupper($rowResultados['paciente'] . " (" . $rowResultados['frecep'] . ", " . $rowResultados['boriginal']) . ")</option>";
                                        }
                                        ?>
                                    </select>
                                </div>
                                <div class='col-md-4'>
                                    <button type='submit' id='submit' name='export' class='btn btn-danger w-100'><i class='fas fa-file-import'></i>&nbsp;&nbsp;Exportar Word</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>

            <div align="center">
                <br>
                <?php echo $MENSAJE; ?>
            </div>

            <div id="footer">
                <div class='col-md-12' align='center'>
                    <a class='btn btn-sm btn-dark' href='https://www.onco.cmw.sld.cu' role='button'>Web ONCO</a>
                </div>
            </div>
        </div>
    </div>
</body>

</html>