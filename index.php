<?php
include("class/security/index.php");
include("conn/conn.php");
require("class/PHPExcel.php");
require("class/fpdf/fpdf.php");

// create document
$pdf = new FPDF();
$pdf->AddPage();

/* Limpiando el directorio de informes anteriores */
$MENSAJE = "";
$DIRECTORIO = "import/";
$HANDLE = opendir($DIRECTORIO);
while ($FILE = readdir($HANDLE)) {
    if ($FILE != "." && $FILE != ".." && $FILE != ".htaccess" && $FILE != ".gitkeep") {
        unlink($DIRECTORIO . $FILE);
    }
}

$col_ano            = array();
$col_noinforme      = array();
$col_boriginal      = array();
$col_organo         = array();
$col_paciente       = array();
$col_hospital       = array();
$col_diagnostico    = array();

$CUENTA_AGREGADOS = 0;
$CUENTA_NO_AGREGADOS = 0;

/* Obteniendo lista de resultados */
$RESULTADOS     = $mysqli->query("SELECT * FROM tbl_biopsias ORDER BY paciente");

/* Procedimiento para importar excel a bd */
if (isset($_POST["import"])) {
    $allowedFileType = ['text/xlsx', 'text/xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    if (in_array($_FILES["file"]["type"], $allowedFileType)) {
        $archivos = 'import/' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $archivos);
        $excel = PHPExcel_IOFactory::load($archivos);
        $excel->setActiveSheetIndex(0);
        $numerofila = $excel->setActiveSheetIndex(0)->getHighestRow();
		$numerofilaok = $numerofila - 2;
        /* Obteniendo los valores de filas en la columna correspondiente */
        for ($i = 0; $i < 26; $i++) {
            $columnLetter = chr($i + 65);
            switch ($excel->getActiveSheet()->getCell($columnLetter . '2')->getValue()) {
                case 'ano':
                    for ($j = 3; $j <= $numerofila; $j++) {
                        $col_ano[$j] = $excel->getActiveSheet()->getCell($columnLetter . $j)->getValue();
                    }
                    break;
                case 'noinforme':
                    for ($k = 3; $k <= $numerofila; $k++) {
                        $col_noinforme[$k] = $excel->getActiveSheet()->getCell($columnLetter . $k)->getValue();
                    }
                    break;
                case 'boriginal':
                    for ($l = 3; $l <= $numerofila; $l++) {
                        $col_boriginal[$l] = $excel->getActiveSheet()->getCell($columnLetter . $l)->getValue();
                    }
                    break;
                case 'organo':
                    for ($m = 3; $m <= $numerofila; $m++) {
                        $col_organo[$m] = $excel->getActiveSheet()->getCell($columnLetter . $m)->getValue();
                    }
                    break;
                case 'paciente':
                    for ($n = 3; $n <= $numerofila; $n++) {
                        $col_paciente[$n] = empty($excel->getActiveSheet()->getCell($columnLetter . $n)->getValue()) ? "SIN NOMBRE DE PACIENTE" : $excel->getActiveSheet()->getCell($columnLetter . $n)->getValue();
                    }
                    break;
                case 'hospital':
                    for ($o = 3; $o <= $numerofila; $o++) {
                        $col_hospital[$o] = empty($excel->getActiveSheet()->getCell($columnLetter . $o)->getValue()) ? "SIN HOSPITAL DEFINIDO" : $excel->getActiveSheet()->getCell($columnLetter . $o)->getValue();
                    }
                    break;
                case 'diagnostico':
                    for ($p = 3; $p <= $numerofila; $p++) {
                        $col_diagnostico[$p] = empty($excel->getActiveSheet()->getCell($columnLetter . $p)->getValue()) ? "SIN DIAGNOSTICO DEFINIDO" : $excel->getActiveSheet()->getCell($columnLetter . $p)->getValue();
                    }
                    break;
            }
        }

        /* Escribiendo en la BD los arrays en el orden correspondiente */
        for ($q = 3; $q <= $numerofila; $q++) {
            $check = mysqli_query($mysqli, "SELECT * FROM tbl_biopsias WHERE ano = '$col_ano[$q]' AND noinforme = '$col_noinforme[$q]' AND boriginal = '$col_boriginal[$q]' AND organo = '$col_organo[$q]' AND paciente = '$col_paciente[$q]' AND hospital = '$col_hospital[$q]' AND diagnostico = '$col_diagnostico[$q]'");

            if (mysqli_num_rows($check) > 0) {
                $CUENTA_NO_AGREGADOS++;
            } else {
                $resultados = mysqli_query($mysqli, "INSERT INTO tbl_biopsias(ano, noinforme, boriginal, organo, paciente, hospital, diagnostico) VALUES('$col_ano[$q]', '$col_noinforme[$q]', '$col_boriginal[$q]', '$col_organo[$q]', '$col_paciente[$q]', '$col_hospital[$q]', '$col_diagnostico[$q]')");
                $CUENTA_AGREGADOS++;
            }
        }
        $MENSAJE = "<div class='alert alert-success alert-dismissible fade show' role='alert'><strong>¡Correcto!</strong>&nbsp;Excel procesado satisfactoriamente: $CUENTA_AGREGADOS registros agregados de $numerofilaok</div>";
    } else {
        $MENSAJE = "<div class='alert alert-warning alert-dismissible fade show' role='alert'><strong>¡Error!</strong>&nbsp;No ha seleccionado un archivo Excel con extensi&oacute;n XLSX. Por favor vuelva a intentarlo</div>";
    }
	$archivos = "";
}

/* Procedimiento para exportar paciente a Word */
if (isset($_POST["export"])) {
    /* Obtenemos valores de las variables */
    $PACIENTE   = mysqli_real_escape_string($mysqli, (strip_tags(strtoupper($_POST["cboPaciente"]), ENT_QUOTES)));
    $pieces = explode("|", $PACIENTE);
    $piece_ano          = $pieces['0'];
    $piece_noinforme    = $pieces['1'];
    $piece_boriginal    = $pieces['2'];
    $piece_organo       = $pieces['3'];
    $piece_paciente     = $pieces['4'];
    $piece_hospital     = strlen($pieces['5']) > 0 ? $pieces['5'] : "-";
    $piece_diagnostico  = $pieces['6'];

	// config document
	$pdf->SetTitle('Nombre');
	$pdf->SetAuthor('AZUfre, Computer Solutions');
	$pdf->SetCreator('FPDF Maker');

	// add image
	$pdf->Image('assets/img/logo.png', null, null, '', '', '','','C');
	$pdf->Ln(10);

	// add encabezado
	$pdf->SetFont('Arial', '', 14);
	$pdf->Cell(0, 10, utf8_decode('INFORME ANATOMOPATOLÓGICO'), 0, 1, 'C');
	$pdf->Ln(15);

	// add title
	$pdf->SetFont('Arial', 'B', 11);
	$pdf->Cell(0, 6, 'Informe enviado desde:', 0, 1);
	$pdf->SetFont('Arial', '', 11);
	$pdf->Cell(0, 5, utf8_decode('Hospital Clínico Quirúrgico Hermanos Ameijeiras'), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('Centro Nacional De Referencia De Anatomía Patológica'), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('San Lázaro 701, La Habana 3, Telef: 78761630'), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('Email: cenrap@hha.sld.uu / telepatol@hha.sld.cu'), 0, 1);

	$pdf->Ln(16);

	$pdf->SetFont('Arial', '', 12);
	$pdf->Cell(0, 5, utf8_decode('BIOPSIA Nro.: ' . "CR" . $piece_ano . $piece_noinforme), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('BIOPSIA ORIGINAL: ' . $piece_boriginal), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('ORGANO: ' . $piece_organo), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('NOMBRE DEL PACIENTE: ' . $piece_paciente), 0, 1);
	$pdf->Cell(0, 5, utf8_decode('HOSPITAL: ' . $piece_hospital), 0, 1);
	$pdf->Ln(4);
	$pdf->Cell(0, 5, utf8_decode('DIAGNOSTICO: '), 0, 1);
	$pdf->MultiCell(0, 5, utf8_decode($piece_diagnostico), 0, 1);

	// output file
	$pdf->Output(utf8_decode($piece_paciente). '.pdf', 'D');
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
    <title>Resultados de Biopsias</title>

    <!-- Enlazando el CSS de Bootstrap -->
    <link href="assets/css/bootstrap.css" rel="stylesheet" media="screen">
    <link href="assets/css/main.css" rel="stylesheet" media="screen">
    <link href="assets/css/fontawesome.css" rel="stylesheet" media="screen">
    <!-- Enlazando el CSS de Bootstrap -->
    <!-- Opcional: enlazando el JavaScript de Bootstrap -->
    <script src="assets/js/jquery-3.6.0.js"></script>
    <script src="assets/js/popper.js"></script>
    <script src="assets/js/bootstrap.js"></script>
    <script src="assets/js/fontawesome.js"></script>
    <!-- Opcional: enlazando el JavaScript de Bootstrap -->
    <script>
        // AUTO HIDE ALERTS
        window.setTimeout(function() {
            $(".alert").fadeTo(500, 0).slideUp(500, function() {
                $(this).remove();
				window.location.href='/';
            });
        }, 2500);
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
                    <div align="center" class="text-success" style="font-size:22px">Importar BD en Excel</div>
                    <div align="center" style="font-size:2px">&nbsp;</div>
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
                <div class="card bg-primary-30 border-primary mb-3">
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center"><i class="fas fa-file-pdf fa-6x text-primary"></i></div>
                    <div align="center" style="font-size:10px">&nbsp;</div>
                    <div align="center" class="text-primary" style="font-size:22px">Exportar resultados a PDF</div>
                    <div align="center" style="font-size:2px">&nbsp;</div>
                    <div class="card-body">
                        <form action='' method='post' name='frmExcelGenerateWord' id='frmExcelGenerateWord' enctype='multipart/form-data'>
                            <div class='row'>
                                <div class='col-md-8'>
                                    <select class="form-select" name="cboPaciente" id="responsive_text" aria-label="Floating label select example" required>
                                        <option disabled value="" selected hidden>Seleccione el Paciente</option>
                                        <?php
                                        while ($rowResultados = $RESULTADOS->fetch_assoc()) {
                                            echo "<option style='white-space:nowrap; text-overflow:elipsis; overflow:hidden;' value='" . $rowResultados['ano'] . "|" . $rowResultados['noinforme'] . "|" . $rowResultados['boriginal'] . "|" . $rowResultados['organo'] . "|" . $rowResultados['paciente'] . "|" . $rowResultados['hospital'] . "|" . $rowResultados['diagnostico'] . "'>" . strtoupper($rowResultados['paciente'] . " (CR" . $rowResultados['ano'] . $rowResultados['noinforme'] . ", " . $rowResultados['boriginal']) . ")</option>";
                                        }
                                        ?>
                                    </select>
                                </div>
                                <div class='col-md-4'>
                                    <button type='submit' id='submit' name='export' class='btn btn-primary w-100'><i class='fas fa-file-import'></i>&nbsp;&nbsp;Exportar PDF</button>
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
                    <a class='btn btn-sm btn-dark' href='class/security/exit.php' role='button'>Web ONCO</a>
                </div>
            </div>
        </div>
    </div>
</body>

</html>