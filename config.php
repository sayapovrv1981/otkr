<?php
// для подключения к бд
error_reporting(E_ALL & ~E_NOTICE);
require_once("utils.php");
require_once("PHPExcel.php");

define('DB_USER', 'root');
define('DB_PASS', '123');
define('DB_HOST', 'localhost');
define('DB_NAME', 'otkr');
define('ROWS_NUMBER_ONPAGE', 3);
define('PAGE_LINK', 3);
define('MAX_FILE_SIZE', '3M');

$db_link = mysqli_connect(DB_HOST, DB_USER, DB_PASS, DB_NAME);

if ($db_link == false){
    print("Ошибка: Невозможно подключиться к MySQL " . mysqli_connect_error());
}
else {
    //print("Соединение установлено успешно");
}
