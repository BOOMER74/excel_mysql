<?php
	assert_options(ASSERT_ACTIVE, 1);
	assert_options(ASSERT_BAIL, 1);

	require_once "PHPExcel.php";

	require_once "../library/excel_mysql.php";

	define("EXCEL_MYSQL_DEBUG", false);

	$connection = new mysqli("localhost", "user", "pass", "excel_mysql_base");

	if ($connection->connect_errno) {
		throw new Exception("MySQL connection error!");
	}

	if (!$connection->set_charset("utf8")) {
		throw new Exception("MySQL set charset error!");
	}

	if (!$connection->query("DROP TABLE IF EXISTS excel_mysql_data, excel_mysql_by_index, excel_mysql_iterate, excel_mysql_by_index_with_option_1, excel_mysql_by_index_with_option_2, excel_mysql_by_index_with_option_3, excel_mysql_by_index_with_option_4, excel_mysql_by_index_with_option_5")) {
		throw new Exception("MySQL query error!");
	}

	if (!$connection->query("CREATE TABLE `excel_mysql_data` (`id` INT(11) NOT NULL AUTO_INCREMENT, `first_name` VARCHAR(50) NOT NULL, `last_name` VARCHAR(50) NOT NULL, `email`   VARCHAR(100) NOT NULL, `pay` FLOAT(10, 2) NOT NULL, PRIMARY KEY (`id`))")) {
		throw new Exception("MySQL query error!");
	}

	if (!$connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('John', 'Smith', 'j.smith@example.com', 10000.00)")) {
		throw new Exception("MySQL query error!");
	}

	if (!$connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('Steve', 'Smith', 's.smith@example.com', 11000.00)")) {
		throw new Exception("MySQL query error!");
	}

	if (!$connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('Oscar', 'Wild', 'o.wild@example.com', 12250.59)")) {
		throw new Exception("MySQL query error!");
	}

	$excel_mysql_import_export = new Excel_mysql($connection, "./test.xlsx");

	assert($excel_mysql_import_export->mysql_to_excel("excel_mysql_data", "Экспорт"));
	assert(file_exists("test.xlsx"));

	assert($excel_mysql_import_export->excel_to_mysql_iterate(array("excel_mysql_iterate")));
	assert($excel_mysql_import_export->excel_to_mysql_by_index("excel_mysql_by_index"));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_1",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		)
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
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
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
		"excel_mysql_by_index_with_option_3",
		0,
		array(
			"id",
			"first_name",
			"last_name",
			"email",
			"pay"
		),
		1
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
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
		array(
			"pay" =>
				function ($value) {
					return floatval($value) > 20000.0;
				}
		)
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
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
		false,
		array(
			"pay" =>
				function ($value) {
					return $value * 2;
				}
		)
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
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
		1
	));

	assert($excel_mysql_import_export->excel_to_mysql_by_index(
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
		false,
		false,
		false,
		array(
			"INT(11) NOT NULL AUTO_INCREMENT",
			"VARCHAR(50) NOT NULL",
			"VARCHAR(50) NOT NULL",
			"VARCHAR(100) NOT NULL",
			"FLOAT(10,2) NOT NULL"
		),
		array("id" => "PRIMARY KEY")
	));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		array(
			"first_name",
			"last_name"
		)
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		array(
			"first_name",
			"last_name"
		),
		array(
			"Имя",
			"Фамилия"
		)
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		false,
		false,
		2
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		false,
		false,
		1,
		1
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		false,
		false,
		false,
		false,
		array(
			"pay" =>
				function ($value) {
					return floatval($value) > 20000.0;
				}
		)
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		false,
		false,
		false,
		false,
		false,
		"pay > 10000"
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		false,
		false,
		false,
		false,
		false,
		false,
		array(
			"pay" =>
				function ($value) {
					return $value . " руб.";
				}
		)
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	assert($excel_mysql_import_export->mysql_to_excel(
		"excel_mysql_data",
		"Экспорт",
		array(
			"id",
			"first_name",
			"last_name",
			"pay"
		),
		false,
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
	));
	assert(file_exists("test.xlsx"));

	if (!unlink("test.xlsx")) {
		throw new Exception("Remove file error!");
	}

	if (!$connection->query("DROP TABLE excel_mysql_data, excel_mysql_by_index, excel_mysql_iterate, excel_mysql_by_index_with_option_1, excel_mysql_by_index_with_option_2, excel_mysql_by_index_with_option_3, excel_mysql_by_index_with_option_4, excel_mysql_by_index_with_option_5")) {
		throw new Exception("MySQL query error!");
	}