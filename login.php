<?php
error_reporting(0);
$MENSAJE = "";

// Inicio del procedimiento de autenticacion
if (isset($_POST['login'])) {
	$ldap_host 		= "172.30.1.2";									// IP de Servidor del dominio
	$ldap_port 		= "389";										// Puerto LDAP del Servidor de dominio
	$ldap_domain 	= "@cmw.smcsalud.cu";							// Dominio de red
	$user 			= $_POST['txtLoginUsername'];					// Nombre de usuario capturado
	$user_full 		= $_POST['txtLoginUsername'] . $ldap_domain;	// Nombre de usuario capturado con dominio
	$pswd 			= $_POST['txtPassUsername'];					// Contraseña capturada
	$base_dn 		= "OU=USUARIOS,DC=cmw,DC=smcsalud,DC=cu";		// Unidad Organizativa de los usuarios del dominio
	$base_group 	= "CN=PARTE,DC=cmw,DC=smcsalud,DC=cu";			// Grupo en el cual se va a buscar al usuario capturado

	if ($fp = @fsockopen($ldap_host, $ldap_port, $ERROR_NO, $ERROR_STR, (float)0.5)) {
		fclose($fp);
		$ldap = ldap_connect($ldap_host);
		ldap_set_option($ldap, LDAP_OPT_PROTOCOL_VERSION, 3);
		ldap_set_option($ldap, LDAP_OPT_REFERRALS, 0);
		$autenticado = ldap_bind($ldap, $user_full, $pswd);

		if ($autenticado) {
			$filter = "(&(objectClass=user) (samaccountname=" . $user . ") (memberOf=" . $base_group . "))";
			$sr = ldap_search($ldap, $base_dn, $filter);

			if (count(ldap_get_entries($ldap, $sr)) == 1) {
				$_SERVER = array();
				$_SESSION = array();
				$_SESSION["autentica"] = "NO";
				$MENSAJE = "<div class='alert alert-danger alert-dismissible fade show' role='alert' id='success-alert'><strong>" . ucfirst(strtolower($user)) . "</strong>,&nbsp;usted no tiene permisos para usar este Sistema<br>Si esto le parece incorrecto contacte al Administrador de la Red</div>";
			} else {
				$attributes = array("displayname");
				$filter = "(&(sAMAccountName=$user))";
				$result = ldap_search($ldap, $base_dn, $filter, $attributes);
				$entries = ldap_get_entries($ldap, $result);

				session_start();
				$_SESSION["user"] = $user;
				$_SESSION["name"] = $entries[0]['displayname'][0];
				$_SESSION["autentica"] = "SI";
				echo "<script>window.location.href='/'; </script>";
			}
		} else {
			$MENSAJE = "<div class='alert alert-warning alert-dismissible fade show' role='alert' id='success-alert'><strong>¡Error!</strong>&nbsp;Usuario/contraseña inv&aacute;lidos o contraseña expirada</div>";
		}
	} else {
		$MENSAJE = "<div class='alert alert-danger alert-dismissible fade show' role='alert' id='success-alert'><strong>¡Error!</strong>&nbsp;No se ha podido conectar al Controlador de Dominio SMC<br>Contacte al Administrador de la Red</div>";
	}
}
// Fin del procedimiento de autenticacion
?>

<!DOCTYPE html>
<html lang="es">

<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">

	<link rel="icon" href="assets/img/favicon.svg">
	<title>Informes Biopsias</title>

	<link href="assets/css/bootstrap.css" rel="stylesheet" media="screen">
	<link href="assets/css/fontawesome.css" rel="stylesheet" media="screen">
	<link href="assets/css/main.css" rel="stylesheet" media="screen">
	<link href="assets/css/signin.css" rel="stylesheet" media="screen">
	<style>
		body {
			padding-top: 0px;
		}
	</style>

	<script src="assets/js/jquery-3.6.4.js"></script>
	<script src="assets/js/bootstrap.js"></script>
	<script src="assets/js/fontawesome.js"></script>
	<script src="assets/js/main.js"></script>
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
	<div class="container">
		<div align="center">
			<div class="row">
				<div class="col" align="center"><i style="font-size:230px" class="fas fa-microscope text-dark"></i></div>
			</div>

			<div class="fs-3 text-dark">Informes Biopsias, autent&iacute;quese</div>
			<div align="center" style="font-size:16px">&nbsp;</div>

			<form name="frmInicio" method="post" action="" class="form-signin">
				<div class="form-floating">
					<input type="text" class="form-control" name="txtLoginUsername" id="txtLoginUsername" placeholder="Usuario" autocomplete="on" required autofocus>
					<label class="text-secondary" for="txtLoginUsername">Usuario</label>
				</div>

				<div align="center" style="font-size:3px">&nbsp;</div>

				<div class="form-floating">
					<input type="password" class="form-control" name="txtPassUsername" id="txtPassUsername" placeholder="Contrase&ntilde;a" autocomplete="off" required>
					<label class="text-secondary" for="txtPassUsername">Contrase&ntilde;a
				</div>

				<div align="center" style="font-size:6px">&nbsp;</div>

				<button class="w-100 btn btn-lg btn-dark" type="submit" name="login">Acceder</button>
			</form>

			<?php echo $MENSAJE ?>
		</div>
	</div>

	<div id="footer">
		<div class='col-md-12' align='center'>
			<a class='btn btn-sm btn-dark' href='https://www.onco.cmw.sld.cu' role='button'>Web ONCO</a>
		</div>
	</div>
</body>

</html>