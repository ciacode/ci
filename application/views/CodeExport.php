<?php

//word
header("Content-Type: application/msword");
header('Content-disposition: inline; filename="exportWord.doc"');

//excel
header('Content-Type: application/xls');
header('Content-Disposition: attachment; filename="exportExcel.xls"');

?>