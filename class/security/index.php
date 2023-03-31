<?php
	@session_start();
	if(isset($_SESSION["autentica"]) != "SI"){
		echo "<script>window.location.href='../../login.php'; </script>";
		exit();
	}
?>
