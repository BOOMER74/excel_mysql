<?php
	// Подключаем библиотеку
	require_once "PHPExcel.php";

	/**
	 * Класс импорта файла Excel в таблицу MySQL и экспорта таблицы MySQL в файл Excel
	 */
	class Excel_mysql {
		/**
		 * @var mysqli - Подключение к базе данных
		 */
		private $mysql_connect;
		/**
		 * @var string - Имя файла для импорта/экспорта
		 */
		private $excel_file;

		/**
		 * Конструктор класса
		 *
		 * @param mysqli $connection - Подключение к базе данных
		 * @param string $filename   - Имя файла для импорта/экспорта
		 */
		function __construct($connection, $filename) {
			$this->mysql_connect = $connection;
			$this->excel_file = $filename;
		}

		/**
		 * Функция преобразования листа Excel в таблицу MySQL, с учетом объединенных строк и столбцов. Значения берутся уже вычисленными
		 *
		 * @param PHPExcel_Worksheet $worksheet                - Лист Excel
		 * @param string             $table_name               - Имя таблицы MySQL
		 * @param int|array          $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n)
		 * @param bool|int           $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array         $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int           $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array         $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11) NOT NULL". Если не указаны, то используется "TEXT NOT NULL"
		 * @param bool|array         $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string             $table_encoding           - Кодировка таблицы MySQL
		 * @param string             $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool
		 */
		private
		function excel_to_mysql($worksheet, $table_name, $columns_names, $start_row_index, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine) {
			// Проверяем соединение с MySQL
			if (!$this->mysql_connect->connect_error) {
				// Строка для названий столбцов таблицы MySQL
				$columns = array();
				// Количество столбцов на листе Excel
				$columns_count = PHPExcel_Cell::columnIndexFromString($worksheet->getHighestColumn());

				// Если в качестве имен столбцов передан массив, то проверяем соответствие его длинны с количеством столбцов
				if ($columns_names) {
					if (is_array($columns_names)) {
						if (count($columns_names) != $columns_count) {
							return false;
						}
					}
				}

				// Если указаны типы столбцов
				if ($table_types) {
					if (is_array($table_types)) {
						// Проверяем количество столбцов и типов
						if (count($table_types) != count($columns_names)) {
							return false;
						}
					}
				}

				// Проверяем, что $columns_names - массив и $unique_column_for_update находиться в его пределах
				if ($unique_column_for_update) {
					$unique_column_for_update = is_array($columns_names) ? ($unique_column_for_update <= count($columns_names) ? "`" . $columns_names[$unique_column_for_update - 1] . "`" : false) : false;
				}

				// Перебираем столбцы листа Excel и генерируем строку с именами через запятую
				for ($column = 0; $column < $columns_count; $column++) {
					/** @noinspection PhpDeprecationInspection */
					$columns[] = "`" . (is_array($columns_names) ? $columns_names[$column] : ($columns_names == 0 ? "column" . $column : $worksheet->getCellByColumnAndRow($column, $columns_names)->getCalculatedValue())) . "`";
				}

				$query_string = "DROP TABLE IF EXISTS `" . $table_name . "`";

				if (defined("EXCEL_MYSQL_DEBUG")) {
					if (EXCEL_MYSQL_DEBUG) {
						var_dump($query_string);
					}
				}

				// Удаляем таблицу MySQL, если она существовала (если не указан столбец с уникальным значением для обновления)
				if ($unique_column_for_update ? true : $this->mysql_connect->query($query_string)) {
					$columns_types = $ignore_columns = array();

					// Обходим столбцы и присваиваем типы
					foreach ($columns as $index => $value) {
						if ($value != "``") {
							if ($table_types) {
								$columns_types[] = $value . " " . $table_types[$index];
							} else {
								$columns_types[] = $value . " TEXT NOT NULL";
							}
						} else {
							$ignore_columns[] = $index;

							unset($columns[$index]);
						}
					}

					// Если указаны ключевые поля, то создаем массив ключей
					if ($table_keys) {
						$columns_keys = array();

						foreach ($table_keys as $key => $value) {
							$columns_keys[] = $key . " (`" . $value . "`)";
						}

						$columns_keys = ", " . implode(", ", $columns_keys);
					} else {
						$columns_keys = "";
					}

					$query_string = "CREATE TABLE IF NOT EXISTS `" . $table_name . "` (" . implode(", ", $columns_types) . $columns_keys . ") COLLATE = '" . $table_encoding . "' ENGINE = " . $table_engine;

					if (defined("EXCEL_MYSQL_DEBUG")) {
						if (EXCEL_MYSQL_DEBUG) {
							var_dump($query_string);
						}
					}

					// Создаем таблицу MySQL
					if ($this->mysql_connect->query($query_string)) {
						// Коллекция значений уникального столбца для удаления несуществующих строк в файле импорта (используется при обновлении)
						$id_list_in_import = array();

						// Количество строк на листе Excel
						$rows_count = $worksheet->getHighestRow();

						// Перебираем строки листа Excel
						for ($row = ($start_row_index ? $start_row_index : (is_array($columns_names) ? 1 : $columns_names + 1)); $row <= $rows_count; $row++) {
							// Строка со значениями всех столбцов в строке листа Excel
							$values = array();

							// Перебираем столбцы листа Excel
							for ($column = 0; $column < $columns_count; $column++) {
								if (in_array($column, $ignore_columns)) {
									continue;
								}

								// Строка со значением объединенных ячеек листа Excel
								$merged_value = "";
								// Ячейка листа Excel
								$cell = $worksheet->getCellByColumnAndRow($column, $row);

								// Перебираем массив объединенных ячеек листа Excel
								foreach ($worksheet->getMergeCells() as $mergedCells) {
									// Если текущая ячейка - объединенная,
									if ($cell->isInRange($mergedCells)) {
										// то вычисляем значение первой объединенной ячейки, и используем её в качестве значения текущей ячейки
										/** @noinspection PhpDeprecationInspection */
										$merged_value = $worksheet->getCell(explode(":", $mergedCells)[0])->getCalculatedValue();

										break;
									}
								}

								/** @noinspection PhpDeprecationInspection */
								$value = strlen($merged_value) == 0 ? $cell->getCalculatedValue() : $merged_value;
								$value = $transform_functions ? (isset($transform_functions[$columns_names[$column]]) ? $transform_functions[$columns_names[$column]]($value) : $value) : $value;

								// Проверяем, что ячейка не объединенная: если нет, то берем ее значение, иначе значение первой объединенной ячейки
								$values[] = "'" . $this->mysql_connect->real_escape_string($value) . "'";
							}

							// Добавляем или проверяем обновлять ли значение
							$add_to_table = $unique_column_for_update ? false : true;

							// Если обновляем
							if ($unique_column_for_update) {
								// Объединяем массивы для простоты работы
								$columns_values = array_combine($columns, $values);

								// Сохраняем уникальное значение
								$id_list_in_import[] = $columns_values[$unique_column_for_update];

								// Создаем условие выборки
								$where = " WHERE " . $unique_column_for_update . " = " . $columns_values[$unique_column_for_update];

								// Удаляем столбец выборки
								unset($columns_values[$unique_column_for_update]);

								$query_string = "SELECT COUNT(*) AS count FROM `" . $table_name . "`" . $where;

								if (defined("EXCEL_MYSQL_DEBUG")) {
									if (EXCEL_MYSQL_DEBUG) {
										var_dump($query_string);
									}
								}

								// Проверяем есть ли запись в таблице
								$count = $this->mysql_connect->query($query_string);
								$count = $count->fetch_assoc();

								// Если есть, то создаем запрос и обновляем
								if (intval($count['count']) != 0) {
									$set = array();

									foreach ($columns_values as $column => $value) {
										$set[] = $column . " = " . $value;
									}

									$query_string = "UPDATE `" . $table_name . "` SET " . implode(", ", $set) . $where;

									if (defined("EXCEL_MYSQL_DEBUG")) {
										if (EXCEL_MYSQL_DEBUG) {
											var_dump($query_string);
										}
									}

									if (!$this->mysql_connect->query($query_string)) {
										return false;
									}
								} else {
									$add_to_table = true;
								}
							}

							// Добавляем строку в таблицу MySQL
							if ($add_to_table) {
								$query_string = "INSERT INTO `" . $table_name . "` (" . implode(", ", $columns) . ") VALUES (" . implode(", ", $values) . ")";

								if (defined("EXCEL_MYSQL_DEBUG")) {
									if (EXCEL_MYSQL_DEBUG) {
										var_dump($query_string);
									}
								}

								if (!$this->mysql_connect->query($query_string)) {
									return false;
								}
							}
						}

						if (!empty($id_list_in_import)) {
							if (defined("EXCEL_MYSQL_DEBUG")) {
								$query_string = "DELETE FROM `" . $table_name . "` WHERE " . $unique_column_for_update . " NOT IN (" . implode(", ", $id_list_in_import) . ")";

								if (EXCEL_MYSQL_DEBUG) {
									var_dump($query_string);
								}
							}

							$this->mysql_connect->query($query_string);
						}

						return true;
					}
				}
			}

			return false;
		}

		/**
		 * Функция импорта листа Excel по индексу
		 *
		 * @param string     $table_name               - Имя таблицы MySQL
		 * @param int        $index                    - Индекс листа Excel
		 * @param int|array  $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n)
		 * @param bool|int   $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int   $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11)"
		 * @param bool|array $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string     $table_encoding           - Кодировка таблицы MySQL
		 * @param string     $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool
		 */
		public
		function excel_to_mysql_by_index($table_name, $index = 0, $columns_names = 0, $start_row_index = false, $transform_functions = false, $unique_column_for_update = false, $table_types = false, $table_keys = false, $table_encoding = "utf8_general_ci", $table_engine = "InnoDB") {
			// Загружаем файл Excel
			$PHPExcel_file = PHPExcel_IOFactory::load($this->excel_file);

			// Выбираем лист Excel
			$PHPExcel_file->setActiveSheetIndex($index);

			return $this->excel_to_mysql($PHPExcel_file->getActiveSheet(), $table_name, $columns_names, $start_row_index, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine);
		}

		/**
		 * Функция импорта всех листов Excel
		 *
		 * @param array      $tables_names             - Массив имен таблиц MySQL
		 * @param int|array  $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n)
		 * @param bool|int   $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int   $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11)"
		 * @param bool|array $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string     $table_encoding           - Кодировка таблицы MySQL
		 * @param string     $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool
		 */
		public
		function excel_to_mysql_iterate($tables_names, $columns_names = 0, $start_row_index = false, $transform_functions = false, $unique_column_for_update = false, $table_types = false, $table_keys = false, $table_encoding = "utf8_general_ci", $table_engine = "InnoDB") {
			// Если массив имен содержит хотя бы 1 запись
			if (count($tables_names) > 0) {
				// Загружаем файл Excel
				$PHPExcel_file = PHPExcel_IOFactory::load($this->excel_file);

				// Перебираем все листы Excel и преобразуем в таблицу MySQL
				foreach ($PHPExcel_file->getWorksheetIterator() as $index => $worksheet) {
					// Имя берётся из массива, если элемент не существует, берем 1й и добавляем индекс
					$table_name = array_key_exists($index, $tables_names) ? $tables_names[$index] : $tables_names[0] . $index;

					if (!$this->excel_to_mysql($worksheet, $table_name, $columns_names, $start_row_index, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine)) {
						return false;
					}
				}

				return true;
			}

			return false;
		}

		/**
		 * Функция экспорта таблицы MySQL в файл Excel. Если файл существует, то его 1й лист будет заменен на экспортируемую таблицу
		 *
		 * @param string $table_name     - Имя таблицы MySQL
		 * @param string $worksheet_name - Имя листа Excel
		 * @param string $file_creator   - Автор документа
		 * @param string $excel_format   - Формат файла Excel
		 *
		 * @return bool
		 */
		public
		function mysql_to_excel($table_name, $worksheet_name, $file_creator = "excel_mysql", $excel_format = "Excel2007") {
			// Проверяем соединение с MySQL
			if (!$this->mysql_connect->connect_error) {
				// Запрос MySQL, возвращающий всё таблицу
				if ($query = $this->mysql_connect->query("SELECT * FROM " . $table_name)) {
					// Если таблица MySQL не пустая
					if ($query->num_rows > 0) {
						// Создаем экземпляр класса PHPExcel
						$phpExcel = new PHPExcel();

						// Задаем лист Excel
						$phpExcel->setActiveSheetIndex(0);
						$worksheet = $phpExcel->getActiveSheet();

						// Задаем имя листа Excel
						$worksheet->setTitle($worksheet_name);

						// Задаем автора (создателя файла)
						$phpExcel->getProperties()->setCreator($file_creator);

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
						$writer->save($this->excel_file);

						return true;
					}
				}
			}

			return false;
		}
	}