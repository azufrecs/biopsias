<?php
	@session_start();
	if(isset($_SESSION["autentica"]) != "SI"){
		echo "<script>window.location.href='https://biopsias.onco.cmw.sld.cu/login.php'; </script>";
		exit();
	}
?>
