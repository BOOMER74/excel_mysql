<?php

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
		 *
		 * @throws Exception - Не найдена библиотека PHPExcel
		 */
		function __construct($connection, $filename) {
			// Если библиотека PHPExcel не подключена
			if (!class_exists("\\PHPExcel")) {
				// Выбрасываем исключение
				throw new \Exception("PHPExcel library required!");
			}

			$this->mysql_connect = $connection;
			$this->excel_file    = $filename;
		}

		/**
		 * Функция преобразования листа Excel в таблицу MySQL, с учетом объединенных строк и столбцов. Значения берутся уже вычисленными
		 *
		 * @param PHPExcel_Worksheet $worksheet                - Лист Excel
		 * @param string             $table_name               - Имя таблицы MySQL
		 * @param int|array          $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n). Если указано больше столбцов, чем на листе Excel, будут использованы значения по умолчанию указанных типов столбцов. Если указано ложное значение (null, false, "", 0, -1...) столбец игнорируется
		 * @param bool|int           $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array         $condition_functions      - Массив функций с условиями добавления строки по значению столбца (столбец => функция)
		 * @param bool|array         $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int           $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array         $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11) NOT NULL". Если не указаны, то используется "TEXT NOT NULL"
		 * @param bool|array         $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string             $table_encoding           - Кодировка таблицы MySQL
		 * @param string             $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool - Флаг, удалось ли выполнить функцию в полном объеме
		 */
		private
		function excel_to_mysql($worksheet, $table_name, $columns_names, $start_row_index, $condition_functions, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine) {
			// Проверяем соединение с MySQL
			if (!$this->mysql_connect->connect_error) {
				// Строка для названий столбцов таблицы MySQL
				$columns = array();

				// Количество столбцов на листе Excel
				$columns_count = \PHPExcel_Cell::columnIndexFromString($worksheet->getHighestColumn());

				// Если в качестве имен столбцов передан массив, то проверяем соответствие его длинны с количеством столбцов
				if ($columns_names) {
					if (is_array($columns_names)) {
						$columns_names_count = count($columns_names);

						if ($columns_names_count < $columns_count) {
							return false;
						} elseif ($columns_names_count > $columns_count) {
							$columns_count = $columns_names_count;
						}
					} else {
						return false;
					}
				}

				// Если указаны типы столбцов
				if ($table_types) {
					if (is_array($table_types)) {
						// Проверяем количество столбцов и типов
						if (count($table_types) != count($columns_names)) {
							return false;
						}
					} else {
						return false;
					}
				}

				$table_name = "`{$table_name}`";

				// Проверяем, что $columns_names - массив и $unique_column_for_update находиться в его пределах
				if ($unique_column_for_update) {
					$unique_column_for_update = is_array($columns_names) ? ($unique_column_for_update <= count($columns_names) ? "`{$columns_names[$unique_column_for_update - 1]}`" : false) : false;
				}

				// Перебираем столбцы листа Excel и генерируем строку с именами через запятую
				for ($column = 0; $column < $columns_count; $column++) {
					$column_name = (is_array($columns_names) ? $columns_names[$column] : ($columns_names == 0 ? "column{$column}" : $worksheet->getCellByColumnAndRow($column, $columns_names)->getValue()));

					$columns[] = $column_name ? "`{$column_name}`" : null;
				}

				$query_string = "DROP TABLE IF EXISTS {$table_name}";

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
						if ($value == null) {
							$ignore_columns[] = $index;

							unset($columns[$index]);
						} else {
							if ($table_types) {
								$columns_types[] = "{$value} {$table_types[$index]}";
							} else {
								$columns_types[] = "{$value} TEXT NOT NULL";
							}
						}
					}

					// Если указаны ключевые поля, то создаем массив ключей
					if ($table_keys) {
						$columns_keys = array();

						foreach ($table_keys as $key => $value) {
							$columns_keys[] = "{$value} (`{$key}`)";
						}

						$columns_keys_list = implode(", ", $columns_keys);

						$columns_keys = ", {$columns_keys_list}";
					} else {
						$columns_keys = null;
					}

					$columns_types_list = implode(", ", $columns_types);

					$query_string = "CREATE TABLE IF NOT EXISTS {$table_name} ({$columns_types_list}{$columns_keys}) COLLATE = '{$table_encoding}' ENGINE = {$table_engine}";

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

						// Получаем массив всех объединенных ячеек
						$all_merged_cells = $worksheet->getMergeCells();

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
								$merged_value = null;

								// Ячейка листа Excel
								$cell = $worksheet->getCellByColumnAndRow($column, $row);

								// Перебираем массив объединенных ячеек листа Excel
								foreach ($all_merged_cells as $merged_cells) {
									// Если текущая ячейка - объединенная,
									if ($cell->isInRange($merged_cells)) {
										// то вычисляем значение первой объединенной ячейки, и используем её в качестве значения текущей ячейки
										$merged_value = explode(":", $merged_cells);

										$merged_value = $worksheet->getCell($merged_value[0])->getValue();

										break;
									}
								}

								// Проверяем, что ячейка не объединенная: если нет, то берем ее значение, иначе значение первой объединенной ячейки
								$value = strlen($merged_value) == 0 ? $cell->getValue() : $merged_value;

								// Если задан массив функций с условиями
								if ($condition_functions) {
									if (isset($condition_functions[$columns_names[$column]])) {
										// Проверяем условие
										if (!$condition_functions[$columns_names[$column]]($value)) {
											break;
										}
									}
								}

								$value = $transform_functions ? (isset($transform_functions[$columns_names[$column]]) ? $transform_functions[$columns_names[$column]]($value) : $value) : $value;

								$values[] = "'{$this->mysql_connect->real_escape_string($value)}'";
							}

							// Если количество столбцов не равно количеству значений, значит строка не прошла проверку
							if ($columns_count - count($ignore_columns) != count($values)) {
								continue;
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
								$where = " WHERE {$unique_column_for_update} = {$columns_values[$unique_column_for_update]}";

								// Удаляем столбец выборки
								unset($columns_values[$unique_column_for_update]);

								$query_string = "SELECT COUNT(*) AS count FROM {$table_name}{$where}";

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
										$set[] = "{$column} = {$value}";
									}

									$set_list = implode(", ", $set);

									$query_string = "UPDATE {$table_name} SET {$set_list}{$where}";

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
								$columns_list = implode(", ", $columns);
								$values_list  = implode(", ", $values);

								$query_string = "INSERT INTO {$table_name} ({$columns_list}) VALUES ({$values_list})";

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
							$id_list = implode(", ", $id_list_in_import);

							$query_string = "DELETE FROM {$table_name} WHERE {$unique_column_for_update} NOT IN ({$id_list})";

							if (defined("EXCEL_MYSQL_DEBUG")) {
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
		 * @param int|array  $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n). Если указано больше столбцов, чем на листе Excel, будут использованы значения по умолчанию указанных типов столбцов. Если указано ложное значение (null, false, "", 0, -1...) столбец игнорируется
		 * @param bool|int   $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array $condition_functions      - Массив функций с условиями добавления строки по значению столбца (столбец => функция)
		 * @param bool|array $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int   $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11)"
		 * @param bool|array $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string     $table_encoding           - Кодировка таблицы MySQL
		 * @param string     $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool - Флаг, удалось ли выполнить функцию в полном объеме
		 */
		public
		function excel_to_mysql_by_index($table_name, $index = 0, $columns_names = 0, $start_row_index = false, $condition_functions = false, $transform_functions = false, $unique_column_for_update = false, $table_types = false, $table_keys = false, $table_encoding = "utf8_general_ci", $table_engine = "InnoDB") {
			// Загружаем файл Excel
			$PHPExcel_file = \PHPExcel_IOFactory::load($this->excel_file);

			// Выбираем лист Excel
			$PHPExcel_file->setActiveSheetIndex($index);

			return $this->excel_to_mysql($PHPExcel_file->getActiveSheet(), $table_name, $columns_names, $start_row_index, $condition_functions, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine);
		}

		/**
		 * Функция импорта всех листов Excel
		 *
		 * @param array      $tables_names             - Массив имен таблиц MySQL
		 * @param int|array  $columns_names            - Строка или массив с именами столбцов таблицы MySQL (0 - имена типа column + n). Если указано больше столбцов чем на листе Excel будут использованы значения по умолчанию
		 * @param bool|int   $start_row_index          - Номер строки, с которой начинается обработка данных (например, если 1 строка шапка таблицы). Нумерация начинается с 1, как в Excel
		 * @param bool|array $condition_functions      - Массив функций с условиями добавления строки по значению столбца (столбец => функция)
		 * @param bool|array $transform_functions      - Массив функций для изменения значения столбца (столбец => функция)
		 * @param bool|int   $unique_column_for_update - Номер столбца с уникальным значением для обновления таблицы. Работает если $columns_names - массив (название столбца берется из него по [$unique_column_for_update - 1])
		 * @param bool|array $table_types              - Типы столбцов таблицы (используется при создании таблицы), в SQL формате - "INT(11)"
		 * @param bool|array $table_keys               - Ключевые поля таблицы (тип => столбец)
		 * @param string     $table_encoding           - Кодировка таблицы MySQL
		 * @param string     $table_engine             - Тип таблицы MySQL
		 *
		 * @return bool - Флаг, удалось ли выполнить функцию в полном объеме
		 */
		public
		function excel_to_mysql_iterate($tables_names, $columns_names = 0, $start_row_index = false, $condition_functions = false, $transform_functions = false, $unique_column_for_update = false, $table_types = false, $table_keys = false, $table_encoding = "utf8_general_ci", $table_engine = "InnoDB") {
			// Если массив имен содержит хотя бы 1 запись
			if (count($tables_names) > 0) {
				// Загружаем файл Excel
				$PHPExcel_file = \PHPExcel_IOFactory::load($this->excel_file);

				// Перебираем все листы Excel и преобразуем в таблицу MySQL
				foreach ($PHPExcel_file->getWorksheetIterator() as $index => $worksheet) {
					// Имя берётся из массива, если элемент не существует, берем 1й и добавляем индекс
					$table_name = array_key_exists($index, $tables_names) ? $tables_names[$index] : "{$tables_names[0]}{$index}";

					if (!$this->excel_to_mysql($worksheet, $table_name, $columns_names, $start_row_index, $condition_functions, $transform_functions, $unique_column_for_update, $table_types, $table_keys, $table_encoding, $table_engine)) {
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
		 * @param string     $table_name          - Имя таблицы MySQL
		 * @param string     $worksheet_name      - Имя листа Excel
		 * @param bool|array $columns_names       - Массив имен столбцов в таблице MySQL
		 * @param bool|array $headers_names       - Массив заголовков для первой строки файла
		 * @param bool|int   $start_row_index     - Стартовая строка в таблице MySQL (SQL запрос - LIMIT x)
		 * @param bool|int   $stop_row_index      - Конечная строка в таблице MySQL (SQL запрос - LIMIT 1, x)
		 * @param bool|array $condition_functions - Массив функций с условиями добавления строк в файл Excel (столбец => функция)
		 * @param bool|array $condition_sql_query - Строка прямого условного SQL запроса ("x = y AND x != z")
		 * @param bool|array $transform_functions - Массив функции для изменения значения столбца (столбец => функция)
		 * @param bool|array $cells_formats       - Массив форматов для ячеек по столбцу (столбец => тип из PHPExcel_Style_NumberFormat)
		 * @param string     $file_creator        - Автор документа
		 * @param string     $excel_format        - Формат файла Excel
		 *
		 * @return bool - Флаг, удалось ли выполнить функцию в полном объеме
		 */
		public
		function mysql_to_excel($table_name, $worksheet_name, $columns_names = false, $headers_names = false, $start_row_index = false, $stop_row_index = false, $condition_functions = false, $condition_sql_query = false, $transform_functions = false, $cells_formats = false, $file_creator = "excel_mysql", $excel_format = "Excel2007") {
			// Проверяем соединение с MySQL
			if (!$this->mysql_connect->connect_error) {
				// Проверяем, что $columns_names это массив
				if ($columns_names) {
					if (!is_array($columns_names)) {
						return false;
					}
				}

				// Проверяем, что $headers_names это массив и его длина соответствует $columns_names
				if ($columns_names && $headers_names) {
					if (is_array($headers_names)) {
						if (count($columns_names) != count($headers_names)) {
							return false;
						}
					} else {
						return false;
					}
				}

				// Проверяем, что $cells_formats это массив и его длина соответствует $columns_names
				if ($columns_names && $cells_formats) {
					if (is_array($cells_formats)) {
						if (count($columns_names) != count($cells_formats)) {
							return false;
						}
					} else {
						return false;
					}
				}

				// Проверяем, если задан $cells_formats, но не задан $columns_names
				if ($cells_formats && !$columns_names) {
					return false;
				}

				$columns_names_list = $columns_names ? implode("`, `", $columns_names) : "*";

				if ($columns_names) {
					$columns_names_list = "`{$columns_names_list}`";
				}

				$condition_sql_query = $condition_sql_query ? " WHERE {$condition_sql_query}" : null;

				// Запрос MySQL, возвращающий таблицу
				$query_string = "SELECT {$columns_names_list} FROM {$table_name}";

				if ($condition_sql_query) {
					$query_string = "{$query_string}{$condition_sql_query}";
				}

				if ($start_row_index || $stop_row_index) {
					$limit_start     = $start_row_index ? intval($start_row_index) : "1";
					$limit_separator = $start_row_index && $stop_row_index ? ", " : null;
					$limit_stop      = $stop_row_index ? intval($stop_row_index) : null;

					$query_string = "{$query_string} LIMIT {$limit_start}{$limit_separator}{$limit_stop}";
				}

				if (defined("EXCEL_MYSQL_DEBUG")) {
					if (EXCEL_MYSQL_DEBUG) {
						var_dump($query_string);
					}
				}

				if ($query = $this->mysql_connect->query($query_string)) {
					// Если таблица MySQL не пустая
					if ($query->num_rows > 0) {
						// Создаем экземпляр класса PHPExcel
						$PHPExcel_instance = new \PHPExcel();

						// Задаем лист Excel
						$PHPExcel_instance->setActiveSheetIndex(0);
						$worksheet = $PHPExcel_instance->getActiveSheet();

						// Задаем имя листа Excel
						$worksheet->setTitle($worksheet_name);

						// Задаем автора (создателя файла)
						$PHPExcel_instance->getProperties()->setCreator($file_creator);

						// Если были заданы заголовки, то записываем их в файл
						if ($headers_names) {
							foreach ($headers_names as $column => $value) {
								$worksheet->setCellValueByColumnAndRow($column, 1, $value);
							}

							// Счетчик строк
							$row = 2;
						} else {
							// Счетчик строк
							$row = 1;
						}

						// Перебираем строки как массив с числовым ключом ([0] => 0)
						while ($rows = $query->fetch_array(2)) {
							$values = array();

							// Перебираем столбцы и пишем в лист Excel
							foreach ($rows as $column => $value) {
								// Если задан массив функций с условиями
								if ($condition_functions) {
									if (isset($condition_functions[$columns_names[$column]])) {
										// Проверяем условие
										if (!$condition_functions[$columns_names[$column]]($value)) {
											break;
										}
									}
								}

								$values[$column] = $transform_functions ? (isset($transform_functions[$columns_names[$column]]) ? $transform_functions[$columns_names[$column]]($value) : $value) : $value;
							}

							// Проверяем, что количество значений равно количеству столбцов
							if (count($values) == count($rows)) {
								foreach ($values as $column => $value) {
									$worksheet->setCellValueByColumnAndRow($column, $row, $value);

									$worksheet->getStyleByColumnAndRow($column, $row)->getNumberFormat()->setFormatCode($cells_formats ? $cells_formats[$columns_names[$column]] : PHPExcel_Style_NumberFormat::FORMAT_GENERAL);
								}

								// Увеличиваем счетчик
								$row++;
							}
						}

						// Создаем "писателя"
						$writer = \PHPExcel_IOFactory::createWriter($PHPExcel_instance, $excel_format);

						// Сохраняем файл
						$writer->save($this->excel_file);

						return true;
					}
				}
			}

			return false;
		}

		/**
		 * Геттер имени файла
		 *
		 * @return string - Имя файла
		 */
		public
		function getFileName() {
			return $this->excel_file;
		}

		/**
		 * Сеттер имени файла
		 *
		 * @param string $filename - Новое имя файла
		 */
		public
		function setFileName($filename) {
			$this->excel_file = $filename;
		}

		/**
		 * Геттер подключения к MySQL
		 *
		 * @return mysqli - Подключение MySQL
		 */
		public
		function getConnection() {
			return $this->mysql_connect;
		}

		/**
		 * Сеттер подключения к MySQL
		 *
		 * @param mysqli $connection - Новое подключение MySQL
		 */
		public
		function setConnection($connection) {
			$this->mysql_connect = $connection;
		}
	}