<?php
// Функция преобразования листа Excel в таблицу MySQL, с учетом объединенных строк и столбцов.
// Значения берутся уже вычисленными. Параметры:
//     $worksheet - лист Excel
//     $connection - соединение с MySQL (mysqli)
//     $table_name - имя таблицы MySQL
//     $columns_name_line - строка с именами столбцов таблицы MySQL (0 - имена типа column + n)
function excel2mysql($worksheet, $connection, $table_name, $columns_name_line = 0,$column_name_str)
{
// Проверяем соединение с MySQL
if (!$connection->connect_error)
  {
    // Строка для названий столбцов таблицы MySQL
    $columns_str = "";
    // Количество столбцов на листе Excel
    $columns_count = count($column_name_str);
    //PHPExcel_Cell::columnIndexFromString($worksheet->getHighestColumn());
    //echo 'count:', $columns_count;
    //print_r($column_name_str);
    // Перебираем столбцы листа Excel и генерируем строку с именами через запятую
    for ($column = 0; $column < $columns_count; $column++)
    {
      $columns_str .= $column_name_str["$column"].",";
      //$columns_str .= ($columns_name_line == 0 ? "column" . $column : $worksheet->getCellByColumnAndRow($column, $columns_name_line)->getCalculatedValue()) . ",";
    }
    // Обрезаем строку, убирая запятую в конце
    $columns_str = substr($columns_str, 0, -1);
  //  echo 'str:',$columns_str;
    // Удаляем таблицу MySQL, если она существовала
    if ($connection->query("DROP TABLE IF EXISTS " . $table_name))
    {
      // Создаем таблицу MySQL
      //устанавливаем кодировку utf8
      if (!mysqli_set_charset($connection, "utf8")) {
    //  printf("Ошибка при загрузке набора символов utf8: %s\n", mysqli_error($connection));
      exit();
  } else {
  //    printf("Текущий набор символов: %s\n", mysqli_character_set_name($connection));
  }
    //  echo "ok<br>";
    //  echo "CREATE TABLE " . $table_name . " (" . str_replace(",", " TEXT NOT NULL,", $columns_str) . " TEXT NOT NULL)";
      if ($connection->query("CREATE TABLE " . $table_name . " (" . str_replace(",", " TEXT NOT NULL,", $columns_str) . " TEXT NOT NULL)"))
      {
        // Количество строк на листе Excel
        $rows_count = $worksheet->getHighestRow();
      //    echo "create";
        // Перебираем строки листа Excel
        for ($row = $columns_name_line + 1; $row <= $rows_count; $row++)
        {
          // Строка со значениями всех столбцов в строке листа Excel
          $value_str = "";

          // Перебираем столбцы листа Excel
          for ($column = 0; $column < $columns_count; $column++)
          {
            // Строка со значением объединенных ячеек листа Excel
            $merged_value = "";
            // Ячейка листа Excel
            $cell = $worksheet->getCellByColumnAndRow($column, $row);

            // Перебираем массив объединенных ячеек листа Excel
            foreach ($worksheet->getMergeCells() as $mergedCells)
            {
              // Если текущая ячейка - объединенная,
              if ($cell->isInRange($mergedCells))
              {
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
          $connection->query("INSERT INTO " . $table_name . " (" . $columns_str . ") VALUES (" . $value_str . ")");
        }
      }
      else// создание таблицы импорта не удалось
      {
        return false;
      }
    }
    else //удаление таблицы не удалось
    {
      return false;
    }
  }
  else // соединение с БД не удалось
  {
    return false;
  }

return true;
}
