<?php
	// Подключаем библиотеку
	require_once "PHPExcel.php";

	// Класс импорта-экспорта Excel в MySQL и наоборот
	class excel_mysql {
		// Соединение с базой
		private $mysqlconnect;
		// Переменная файла Excel
		private $excelfile;

		// Конструктор класса
		function __construct($connection, $filename) {
			// Соединение с базой с использованием внешнего соединения
			$this->mysqlconnect = $connection;
			// Имя файла Excel
			$this->excelfile = $filename;
		}

		// Функция преобразования листа Excel в таблицу MySQL, с учетом объединенных строк и столбцов.
		// Значения берутся уже вычисленными. Параметры:
		//   $worksheet - лист Excel
		//   $table_name - имя таблицы MySQL
		//   $columns_name_line - строка с именами столбцов таблицы MySQL (0 - имена типа column + n)
		private function excel2mysql($worksheet, $table_name, $columns_name_line) {
			// Проверяем соединение с MySQL
			if (!$this->mysqlconnect->connect_error) {
				// Строка для названий столбцов таблицы MySQL
				$columns_str = "";
				// Количество столбцов на листе Excel
				$columns_count = PHPExcel_Cell::columnIndexFromString($worksheet->getHighestColumn());

				// Перебираем столбцы листа Excel и генерируем строку с именами через запятую
				for ($column = 0; $column < $columns_count; $column++) {
					$columns_str .= ($columns_name_line == 0 ? "column" . $column : $worksheet->getCellByColumnAndRow($column, $columns_name_line)->getCalculatedValue()) . ",";
				}

				// Обрезаем строку, убирая запятую в конце
				$columns_str = substr($columns_str, 0, -1);

				// Удаляем таблицу MySQL, если она существовала
				if ($this->mysqlconnect->query("DROP TABLE IF EXISTS " . $table_name)) {
					// Создаем таблицу MySQL
					if ($this->mysqlconnect->query("CREATE TABLE " . $table_name . " (" . str_replace(",", " TEXT NOT NULL,", $columns_str) . " TEXT NOT NULL)")) {
						// Количество строк на листе Excel
						$rows_count = $worksheet->getHighestRow();
						// Перебираем строки листа Excel
						for ($row = $columns_name_line + 1; $row <= $rows_count; $row++) {
							// Строка со значениями всех столбцов в строке листа Excel
							$value_str = "";

							// Перебираем столбцы листа Excel
							for ($column = 0; $column < $columns_count; $column++) {
								// Строка со значением объединенных ячеек листа Excel
								$merged_value = "";
								// Ячейка листа Excel
								$cell = $worksheet->getCellByColumnAndRow($column, $row);

								// Перебираем массив объединенных ячеек листа Excel
								foreach ($worksheet->getMergeCells() as $mergedCells) {
									// Если текущая ячейка - объединенная,
									if ($cell->isInRange($mergedCells)) {
										// то вычисляем значение первой объединенной ячейки, и используем её в качестве значения
										// текущей ячейки
										$merged_value = $worksheet->getCell(explode(":", $mergedCells)[0])->getCalculatedValue();
										break;
									}
								}

								// Проверяем, что ячейка не объединенная: если нет, то берем ее значение, иначе значение первой
								// объединенной ячейки
								$value_str .= "'" . (strlen($merged_value) == 0 ? $cell->getCalculatedValue() : $merged_value) . "',";
							}

							// Обрезаем строку, убирая запятую в конце
							$value_str = substr($value_str, 0, -1);

							// Добавляем строку в таблицу MySQL
							$this->mysqlconnect->query("INSERT INTO " . $table_name . " (" . $columns_str . ") VALUES (" . $value_str . ")");
						}
					} else {
						return false;
					}
				} else {
					return false;
				}
			} else {
				return false;
			}

			return true;
		}

		// Импорт листа Excel по индексу. Параметры:
		//   $table_name - имя таблицы MySQL
		//   $index - индекс листа Excel
		//   $columns_name_line - строка с именами столбцов таблицы MySQL (0 - имена типа column + n)
		public function excel2mysql_byindex($table_name, $index = 0, $columns_name_line = 0) {
			// Загружаем файл Excel
			$PHPExcel_file = PHPExcel_IOFactory::load($this->excelfile);

			// Выбираем лист Excel
			$PHPExcel_file->setActiveSheetIndex($index);

			return $this->excel2mysql($PHPExcel_file->getActiveSheet(), $table_name, $columns_name_line);
		}

		// Импорт всех листов Excel. Параметры:
		//   $tables_names - массив имен таблиц MySQL
		//   $columns_name_line - строка с именами столбцов таблицы MySQL (0 - имена типа column + n)
		public function excel2mysql_iterate($tables_names, $columns_name_line = 0) {
			// Если массив имен содержит хотя бы 1 запись
			if (count($tables_names) > 0) {
				// Загружаем файл Excel
				$PHPExcel_file = PHPExcel_IOFactory::load($this->excelfile);

				// Перебираем все листы Excel и преобразуем в таблицу MySQL
				foreach ($PHPExcel_file->getWorksheetIterator() as $index => $worksheet) {
					// Имя берётся из массива, если элемент не существует, берем 1й и добавляем индекс
					$table_name = array_key_exists($index, $tables_names) ? $tables_names[$index] : $tables_names[0] . $index;

					if (!$this->excel2mysql($worksheet, $table_name, $columns_name_line)) {
						return false;
					}
				}
			} else {
				return false;
			}

			return true;
		}

		public function mysql2excel($table_name, $worksheet_name, $excel_format = "Excel2007") {
			// Проверяем соединение с MySQL
			if (!$this->mysqlconnect->connect_error) {
				// Запрос MySQL, возвращающий всё таблицу
				if ($query = $this->mysqlconnect->query("SELECT * FROM " . $table_name)) {
					// Если таблица MySQL не пустая
					if ($query->num_rows > 0) {
						// Создаем экземпляр класса PHPExcel
						$phpExcel = new PHPExcel();

						// Задаем лист Excel
						$phpExcel->setActiveSheetIndex(0);
						$worksheet = $phpExcel->getActiveSheet();

						// Задаем имя листа Excel
						$worksheet->setTitle($worksheet_name);

						// Счетчик строк
						$row = 1;

						// Перебираем строки как массив с числовым ключом ([0] => 0)
						while ($rows = $query->fetch_array(2)) {
							// Перебираем столбцы и пишем в лист Excel
							foreach ($rows as $column => $value) {
								$worksheet->setCellValueByColumnAndRow($column, $row, $value);
							}

							// Увеличиваем счетчик
							$row++;
						}

						// Создаем "писателя"
						$writer = PHPExcel_IOFactory::createWriter($phpExcel, $excel_format);
						// Сохраняем файл
						$writer->save($this->excelfile);
					} else {
						return false;
					}
				} else {
					return false;
				}
			} else {
				return false;
			}

			return true;
		}
	}
?>