<?php
// Process COBList 20170513
include ('PAI_coblist.class.php');
register_shutdown_function('shutDownFunction');

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

function shutDownFunction() { 
    $error = error_get_last();
    // fatal error, E_ERROR === 1
    if ($error['type'] === E_ERROR) { 
        //do your stuff
		echo "Program failed! Please try again using left menu Run COBList. If it keeps failing notify Chris Barlow.";
    } 
}
?>