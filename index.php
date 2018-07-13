<?php
spl_autoload_register(function($class){
    include_once('./'.$class.'.php');
});

$dsn = 'mysql:host=localhost;dbname=test';
$username = 'root';
$password = 'root';
$dbName = 'test';

(new DataDict())->createDataDictExcel($dsn, $username, $password, $dbName);

?>