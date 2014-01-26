<?php
	// Подключаем модуль
	require_once "../excel_mysql.php";

	// Соединение с базой MySQL
	$connection = new mysqli("localhost", "user", "pass", "base");
	// Выбираем кодировку UTF-8
	$connection->set_charset("utf8");

	// Создаем экземпляр класса excel_mysql
	$excel_mysql_import_export = new excel_mysql($connection, "./example.xlsx");

	// Экспортируем таблицу MySQL в Excel
	echo $excel_mysql_import_export->mysql2excel("excel_mysql", "Экспорт") ? "OK\n" : "FAIL\n";

	// Перебираем все листы Excel и преобразуем в таблицу MySQL
	echo $excel_mysql_import_export->excel2mysql_iterate(array("excel_mysql")) ? "OK\n" : "FAIL\n";

	// Преобразуем первый лист Excel в таблицу MySQL
	echo $excel_mysql_import_export->excel2mysql_byindex("excel_mysql_first") ? "OK\n" : "FAIL\n";