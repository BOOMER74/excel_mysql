<?php
	require_once __DIR__ . "/../vendor/autoload.php";
	require_once __DIR__ . "/../library/excel_mysql.php";
	require_once __DIR__ . "/../PHPExcel/Classes/PHPExcel.php";

	define("EXCEL_MYSQL_DEBUG", false);

	class Excel_mysql_test extends PHPUnit_Framework_TestCase {
		public
		function testAll() {
			$connection = new mysqli("localhost", "user", "pass", "excel_mysql_base");

			$this->assertEquals($connection->connect_errno, 0);

			$this->assertTrue($connection->set_charset("utf8"));

			$this->assertTrue($connection->query("DROP TABLE IF EXISTS excel_mysql_data, excel_mysql_by_index, excel_mysql_iterate, excel_mysql_by_index_with_option_1, excel_mysql_by_index_with_option_2, excel_mysql_by_index_with_option_3, excel_mysql_by_index_with_option_4, excel_mysql_by_index_with_option_5, excel_mysql_by_index_with_option_6"));

			$this->assertTrue($connection->query("CREATE TABLE `excel_mysql_data` (`id` INT(11) NOT NULL AUTO_INCREMENT, `first_name` VARCHAR(50) NOT NULL, `last_name` VARCHAR(50) NOT NULL, `email`   VARCHAR(100) NOT NULL, `pay` FLOAT(10, 2) NOT NULL, PRIMARY KEY (`id`))"));

			$this->assertTrue($connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('John', 'Smith', 'j.smith@example.com', 10000.00)"));
			$this->assertTrue($connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('Steve', 'Smith', 's.smith@example.com', 11000.00)"));
			$this->assertTrue($connection->query("INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`) VALUES ('Oscar', 'Wild', 'o.wild@example.com', 12250.59)"));

			$excel_mysql_import_export = new Excel_mysql($connection, "./test.xlsx");

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel("excel_mysql_data", "Экспорт"));
			$this->assertTrue(file_exists("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_iterate(array("excel_mysql_iterate")));
			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index("excel_mysql_by_index"));

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
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

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
				"excel_mysql_by_index_with_option_6",
				0,
				array(
					"id",
					null,
					"last_name",
					"email",
					"pay"
				)
			));

			$this->assertTrue($excel_mysql_import_export->excel_to_mysql_by_index(
				"excel_mysql_by_index_with_option_6",
				0,
				array(
					"id",
					null,
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
					null,
					"VARCHAR(50) NOT NULL",
					"VARCHAR(100) NOT NULL",
					"FLOAT(10,2) NOT NULL"
				),
				array("id" => "PRIMARY KEY")
			));

			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
				"excel_mysql_data",
				"Экспорт",
				array(
					"first_name",
					"last_name"
				)
			));

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
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

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
				"excel_mysql_data",
				"Экспорт",
				false,
				false,
				2
			));

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
				"excel_mysql_data",
				"Экспорт",
				false,
				false,
				1,
				1
			));

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
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

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
				"excel_mysql_data",
				"Экспорт",
				false,
				false,
				false,
				false,
				false,
				"pay > 10000"
			));

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
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
							return "{$value} руб.";
						}
				)
			));

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($excel_mysql_import_export->mysql_to_excel(
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

			$this->assertTrue(file_exists("test.xlsx"));
			$this->assertTrue(unlink("test.xlsx"));

			$this->assertTrue($connection->query("DROP TABLE excel_mysql_data, excel_mysql_by_index, excel_mysql_iterate, excel_mysql_by_index_with_option_1, excel_mysql_by_index_with_option_2, excel_mysql_by_index_with_option_3, excel_mysql_by_index_with_option_4, excel_mysql_by_index_with_option_5, excel_mysql_by_index_with_option_6"));
		}
	}