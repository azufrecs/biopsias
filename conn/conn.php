<?php
$mysqli = new mysqli('127.0.0.1', 'biopsias', 'biopsias2012*/', 'biopsias');
if ($mysqli->connect_error) {
	die('Error : (' . $mysqli->connect_errno . ') ' . $mysqli->connect_error);
}
