<?php
// Process COBList 20170506
include ('PAI_coblist.class.php');
$mCOB = new COBList();
if ($mCOB->Checkfile($msg)) {
	$mCOB->showInfo = $_POST['show'];
	$mCOB->fullRun = $_POST['run'];
	if ($mCOB->ProcessFile($msg)) {
	} else {
		echo $msg;
	}
} else {
	echo $msg;
}
unset($mCOB);
?>