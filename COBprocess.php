<?php
// Process COBList 20180325
include ('PAI_coblist.class.php');
register_shutdown_function('shutDownFunction');

$mCOB = new COBList();
if ($mCOB->Checkfile($msg)) {
	$mCOB->logging = false;
	$mCOB->showInfo = isset($_POST['showinfo']);
	$mCOB->fullRun = isset($_POST['fullrun']);
	if ($mCOB->ProcessFile($msg)) {
	} else {
	error_log ($_SERVER['REMOTE_ADDR'] . '=' . $msg,0);
	echo $msg;
	}
} else {
	error_log ($_SERVER['REMOTE_ADDR'] . '=' . $msg,0);
	echo $msg;
}
unset($mCOB);

function shutDownFunction() { 
    $error = error_get_last();
    // fatal error, E_ERROR === 1
    if ($error['type'] === E_ERROR) { 
        //do your stuff
		error_log ($_SERVER['REMOTE_ADDR'] . '=' . $msg,0);
		echo "Program failed! Please try again using left menu Run COBList. If it keeps failing notify Chris Barlow.";
    } 
}
?>