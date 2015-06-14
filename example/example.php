<?php
	// Подключаем библиотеку
	require_once "PHPExcel.php";
	// Подключаем модуль
	require_once __DIR__ . "/../library/excel_mysql.php";

	// Определяем константу для включения режима отладки (режим отладки выключен)
	define("EXCEL_MYSQL_DEBUG", false);

	// Соединение с базой MySQL
	$connection = new mysqli("localhost", "user", "pass", "excel_mysql_base");
	// Выбираем кодировку UTF-8
	$connection->set_charset("utf8");

	// Создаем экземпляр класса excel_mysql
	$excel_mysql_import_export = new Excel_mysql($connection, "./example.xlsx");

	// Примеры без дополнительных настроек

	// Экспортируем таблицу MySQL в Excel
	echo $excel_mysql_import_export->mysql_to_excel("excel_mysql_data", "Экспорт") ? "OK\n" : "FAIL\n";

	// Перебираем все листы Excel и преобразуем в таблицу MySQL
	echo $excel_mysql_import_export->excel_to_mysql_iterate(array("excel_mysql_iterate")) ? "OK\n" : "FAIL\n";

	// Преобразуем первый лист Excel в таблицу MySQL
	echo $excel_mysql_import_export->excel_to_mysql_by_index("excel_mysql_by_index") ? "OK\n" : "FAIL\n";

	// Примеры с дополнительными настройками

	// Указываем названия столбцов в таблице MySQL
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_1",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		)
	) ? "OK\n" : "FAIL\n";

	// Указываем названия столбцов в таблице MySQL. Дополнительно указываем столбцы которые, в случае отсутствия, будут заполнены значением по умолчанию
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_2",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay",
			"empty_1",
			"empty_2",
			"empty_3"
		)
	) ? "OK\n" : "FAIL\n";

	// Указываем названия столбцов в таблице MySQL и функцию изменения значения для конкретного столбца (например для преобразования дат из Excel в MySQL)
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_3",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		),
		false,
		false,
		array(
			"pay" =>
				function ($value) {
					return $value * 2;
				}
		)
	) ? "OK\n" : "FAIL\n";

	// Экспортируем таблицу MySQL в Excel
	echo $excel_mysql_import_export->mysql_to_excel("excel_mysql_by_index_with_option_3", "Экспорт") ? "OK\n" : "FAIL\n";

	// Указываем названия столбцов в таблице MySQL и уникальный столбец для обновления таблицы
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_1",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		),
		1
	) ? "OK\n" : "FAIL\n";

	// Указываем названия столбцов в таблице MySQL, их типы и ключевое поле
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_4",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		),
		false,
		false,
		false,
		1,
		array(
			"INT(11) NOT NULL AUTO_INCREMENT",
			"VARCHAR(50) NOT NULL",
			"VARCHAR(50) NOT NULL",
			"VARCHAR(100) NOT NULL",
			"FLOAT(10,2) NOT NULL"
		),
		array("id" => "PRIMARY KEY")
	) ? "OK\n" : "FAIL\n";

	// Указываем названия столбцов в таблице MySQL и условия добавления
	echo $excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_5",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		),
		false,
		array(
			"pay" =>
				function ($value) {
					return floatval($value) > 20000.0;
				}
		)
	) ? "OK\n" : "FAIL\n";

	// Изменяем имя файла
	$excel_mysql_import_export->setFileName("export1.xlsx");

	// Экспортируем таблицу MySQL в Excel с указанием какие столбцы выгружать и заголовками столбцов
	echo $excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_by_index_with_option_1",
		"Экспорт",
		array(
			"first_name",
			"last_name"
		),
		array(
			"Имя",
			"Фамилия"
		)
	) ? "OK\n" : "FAIL\n";

	// Изменяем имя файла
	$excel_mysql_import_export->setFileName("export2.xlsx");

	// Экспортируем таблицу MySQL в Excel с указанием какие столбцы выгружать и заголовками столбцов, условиями выборки и преобразованием значения столбца
	echo $excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_by_index_with_option_3",
		"Экспорт",
		array(
			"id",
			"first_name",
			"last_name",
			"pay"
		),
		array(
			"№",
			"Имя",
			"Фамилия",
			"Зарплата"
		),
		false,
		false,
		array(
			"pay" =>
				function ($value) {
					return floatval($value) > 20000.0;
				}
		),
		false,
		array(
			"pay" =>
				function ($value) {
					return "{$value} руб.";
				}
		)
	) ? "OK\n" : "FAIL\n";

	// Изменяем имя файла
	$excel_mysql_import_export->setFileName("export3.xlsx");

	// Экспортируем таблицу MySQL в Excel с указанием какие столбцы выгружать, заголовками столбцов и форматами ячеек
	echo $excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_by_index_with_option_4",
		"Экспорт",
		array(
			"id",
			"first_name",
			"last_name",
			"pay"
		),
		array(
			"№",
			"Имя",
			"Фамилия",
			"Зарплата"
		),
		false,
		false,
		false,
		false,
		false,
		array(
			"id"         => PHPExcel_Style_NumberFormat::FORMAT_NUMBER,
			"first_name" => PHPExcel_Style_NumberFormat::FORMAT_TEXT,
			"last_name"  => PHPExcel_Style_NumberFormat::FORMAT_TEXT,
			"pay"        => PHPExcel_Style_NumberFormat::FORMAT_NUMBER
		)
	) ? "OK\n" : "FAIL\n";